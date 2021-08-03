VERSION 5.00
Begin VB.Form managementdbsaveDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Management Batabase Backup"
   ClientHeight    =   2197
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   3653
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2197
   ScaleWidth      =   3653
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox filenametext 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   351
      TabIndex        =   1
      Top             =   468
      Width           =   2938
   End
   Begin VB.CommandButton TakeBackup 
      BackColor       =   &H8000000D&
      Caption         =   "Take Backup"
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
      Left            =   1053
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   936
      Width           =   1534
   End
   Begin VB.Label filenamelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter File Name:"
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
      Left            =   1170
      TabIndex        =   3
      Top             =   234
      Width           =   1651
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: All backup are saved in Project Directory"
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
      Height          =   481
      Left            =   117
      TabIndex        =   2
      Top             =   1638
      Width           =   3406
   End
End
Attribute VB_Name = "managementdbsaveDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fileSys As New FileSystemObject

Private Sub Form_Load()
With fileSys
   If .FolderExists(App.Path & "\Backup_DB") = False Then .CreateFolder (App.Path & "\Backup_DB")
 End With
End Sub

Private Sub TakeBackup_Click()
If filenametext <> "" Then
   fileSys.CopyFile App.Path & "\Database\admin_data.mdb", App.Path & "\Backup_DB\" & filenametext.Text & "_Management.mdb", True  'want to have a filename as celeste..haha.. you can change it into diff. names ok!hahah
   MsgBox "Yes! Backup Successful! ", vbInformation, "Backup Successful!"
   Unload Me
 Else
   MsgBox "Invalid filename", vbCritical, "Backup Failed!"
 End If
End Sub

