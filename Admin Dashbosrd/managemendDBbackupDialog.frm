VERSION 5.00
Begin VB.Form studentDBbackupDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Database Backup"
   ClientHeight    =   2015
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   3406
   Icon            =   "managemendDBbackupDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2015
   ScaleWidth      =   3406
   ShowInTaskbar   =   0   'False
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
      Left            =   936
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   819
      Width           =   1534
   End
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
      Left            =   234
      TabIndex        =   1
      Top             =   351
      Width           =   2938
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
      Left            =   0
      TabIndex        =   3
      Top             =   1521
      Width           =   3406
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
      Left            =   1053
      TabIndex        =   0
      Top             =   117
      Width           =   1651
   End
End
Attribute VB_Name = "studentDBbackupDialog"
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
   fileSys.CopyFile App.Path & "\Database\user_data.mdb", App.Path & "\Backup_DB\" & filenametext.Text & "_Student.mdb", True  'want to have a filename as celeste..haha.. you can change it into diff. names ok!hahah
   MsgBox "Yes! Backup Successful! ", vbInformation, "Backup Successful!"
   Unload Me
 Else
   MsgBox "Invalid filename", vbCritical, "Backup Failed!"
 End If
End Sub
