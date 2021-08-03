VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FeeStatusDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fee Status"
   ClientHeight    =   4186
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6149
   Icon            =   "FeeStatusDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4186
   ScaleWidth      =   6149
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox transno 
      Alignment       =   2  'Center
      DataField       =   "transid"
      DataSource      =   "submittransadb"
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
      Left            =   3861
      TabIndex        =   11
      Top             =   2574
      Width           =   2002
   End
   Begin MSAdodcLib.Adodc submittransadb 
      Height          =   299
      Left            =   2574
      Top             =   1404
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
      Connect         =   $"FeeStatusDialog.frx":2FACA
      OLEDBString     =   $"FeeStatusDialog.frx":2FB55
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from feetrans"
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
   Begin VB.TextBox Textrn 
      Alignment       =   2  'Center
      DataField       =   "rollno"
      DataSource      =   "submittransadb"
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
      Left            =   3861
      TabIndex        =   8
      Top             =   1872
      Width           =   2002
   End
   Begin VB.CommandButton searchtransresult 
      BackColor       =   &H8000000D&
      Caption         =   "Submit"
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
      Left            =   4212
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3042
      Width           =   1183
   End
   Begin VB.TextBox feerollno 
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
      Left            =   117
      TabIndex        =   2
      Top             =   2106
      Width           =   2353
   End
   Begin VB.CommandButton searchfeeresult 
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
      Left            =   468
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2925
      Width           =   1651
   End
   Begin MSAdodcLib.Adodc matctdetailsado 
      Height          =   364
      Left            =   2457
      Top             =   2691
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
      Connect         =   $"FeeStatusDialog.frx":2FBE0
      OLEDBString     =   $"FeeStatusDialog.frx":2FC6B
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
   Begin VB.Label Label6 
      Height          =   247
      Left            =   2925
      TabIndex        =   12
      Top             =   2223
      Visible         =   0   'False
      Width           =   598
   End
   Begin VB.Label datecap 
      DataField       =   "datee"
      DataSource      =   "submittransadb"
      Height          =   247
      Left            =   2691
      TabIndex        =   10
      Top             =   3159
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label5 
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
      Left            =   4212
      TabIndex        =   9
      Top             =   1638
      Width           =   1183
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"FeeStatusDialog.frx":2FCF6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   598
      Left            =   117
      TabIndex        =   7
      Top             =   3627
      Width           =   5980
   End
   Begin VB.Line Line1 
      X1              =   2925
      X2              =   2925
      Y1              =   1638
      Y2              =   3159
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Transaction ID"
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
      Left            =   4095
      TabIndex        =   5
      Top             =   2340
      Width           =   1651
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUBMIT TRANS. ID"
      BeginProperty Font 
         Name            =   "PanRoman"
         Size            =   10.87
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   247
      Left            =   3393
      TabIndex        =   4
      Top             =   1404
      Width           =   2938
   End
   Begin VB.Label Label2RN 
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
      Left            =   585
      TabIndex        =   3
      Top             =   1755
      Width           =   1183
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK FEE STATUS"
      BeginProperty Font 
         Name            =   "PanRoman"
         Size            =   10.87
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   247
      Left            =   -117
      TabIndex        =   0
      Top             =   1404
      Width           =   2704
   End
   Begin VB.Image Image1 
      Height          =   1638
      Left            =   0
      Picture         =   "FeeStatusDialog.frx":2FDA7
      Stretch         =   -1  'True
      Top             =   -234
      Width           =   6097
   End
End
Attribute VB_Name = "FeeStatusDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub datecap_Click()
datecap.Caption = Date
End Sub

Private Sub Form_Load()
datecap.Caption = Date
submittransadb.Recordset.AddNew
datecap.Caption = Date
Label6.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub Label6_Click()
Label6.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub searchfeeresult_Click()
matctdetailsado.RecordSource = "select * from user_data where username='" + Label6.Caption + "' and rollno='" + feerollno.Text + "'"
matctdetailsado.Refresh


DataEnvironment4feestatus.rsfeestatusresult.Open "select * from feestatus where rollno= '" + feerollno.Text + "'"
feestatusDataReport.Refresh
If DataEnvironment4feestatus.rsfeestatusresult.EOF Then
MsgBox "Roll. No.Doesn't Match, Please Try Again", vbCritical, "Massage"
DataEnvironment4feestatus.rsfeestatusresult.Close
FeeStatusDialog.Show

ElseIf matctdetailsado.Recordset.EOF Then
 MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
 DataEnvironment4feestatus.rsfeestatusresult.Close
    feerollno.SetFocus
    SendKeys "{Home}+{End}"
    
Else
FeeStatusDialog.Visible = False
feestatusDataReport.Show
DataEnvironment4feestatus.rsfeestatusresult.Close

End If

End Sub

Private Sub searchtransresult_Click()
matctdetailsado.RecordSource = "select * from user_data where username='" + Label6.Caption + "' and rollno='" + Textrn.Text + "'"
matctdetailsado.Refresh


If Textrn = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
    Textrn.SetFocus
    SendKeys "{Home}+{End}"

ElseIf transno = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
    transno.SetFocus
    SendKeys "{Home}+{End}"

ElseIf matctdetailsado.Recordset.EOF Then
 MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
    Textrn.SetFocus
    SendKeys "{Home}+{End}"
    
Else
submittransadb.Recordset.AddNew
MsgBox "Successfully Submited!", vbOKOnly + vbInformation, "Success"
datecap.Caption = Date
End If

End Sub

