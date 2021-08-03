VERSION 5.00
Begin VB.Form studentroutineDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routine"
   ClientHeight    =   3731
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   5850
   Icon            =   "studentroutineDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3731
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox chooseyearstudent 
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
      Left            =   1632
      TabIndex        =   5
      Text            =   "Choose Year"
      Top             =   2691
      Width           =   2587
   End
   Begin VB.ComboBox choosecoursestudent 
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
      Left            =   1632
      TabIndex        =   4
      Text            =   "Choose Course"
      Top             =   1989
      Width           =   2587
   End
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
      Left            =   2055
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3042
      Width           =   1651
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   2334
      TabIndex        =   3
      Top             =   2340
      Width           =   1066
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   2217
      TabIndex        =   2
      Top             =   1638
      Width           =   1534
   End
   Begin VB.Line Line1 
      X1              =   1287
      X2              =   4446
      Y1              =   1521
      Y2              =   1521
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ROUTINE"
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
      Left            =   1476
      TabIndex        =   0
      Top             =   1170
      Width           =   2821
   End
   Begin VB.Image Image1 
      Height          =   1521
      Left            =   247
      Picture         =   "studentroutineDialog.frx":2FACA
      Stretch         =   -1  'True
      Top             =   -234
      Width           =   5278
   End
End
Attribute VB_Name = "studentroutineDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
choosecoursestudent.AddItem "BCA"
choosecoursestudent.AddItem "BBM"
chooseyearstudent.AddItem "1 Year"
chooseyearstudent.AddItem "3 Year"
End Sub

Private Sub searchattendence_Click()
DataEnvironment3routine.rsroutinegen.Open "select * from routin_notification where course= '" + choosecoursestudent.Text + "' and batchyear='" + chooseyearstudent.Text + "'"
routineReport.Refresh
If DataEnvironment3routine.rsroutinegen.EOF Then
MsgBox "Selected Data Doesn't Match, Please Try Again", vbCritical, "Massage"
'studentroutineDialog.Visible = False
DataEnvironment3routine.rsroutinegen.Close
Else
studentroutineDialog.Visible = False
routineReport.Show
DataEnvironment3routine.rsroutinegen.Close

End If

End Sub
