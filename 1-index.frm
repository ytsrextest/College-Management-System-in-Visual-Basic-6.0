VERSION 5.00
Begin VB.Form Form2_loginarea 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ABC College | MANAGEMENT/STUDENT SYSTEM- 2021"
   ClientHeight    =   6370
   ClientLeft      =   39
   ClientTop       =   299
   ClientWidth     =   8086
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "1-index.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6370
   ScaleWidth      =   8086
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6565
      Left            =   0
      Picture         =   "1-index.frx":2FACA
      ScaleHeight     =   6513
      ScaleWidth      =   8151
      TabIndex        =   0
      Top             =   -234
      Width           =   8203
      Begin VB.CommandButton exit 
         BackColor       =   &H00FFFFFF&
         Height          =   1066
         Left            =   5824
         MaskColor       =   &H00FFFFFF&
         Picture         =   "1-index.frx":33F2C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5148
         Width           =   1885
      End
      Begin VB.CommandButton helpbtn 
         BackColor       =   &H00FFFFFF&
         Height          =   1066
         Left            =   377
         MaskColor       =   &H00FFFFFF&
         Picture         =   "1-index.frx":34767
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5148
         Width           =   2002
      End
      Begin VB.CommandButton Contactbtn 
         BackColor       =   &H00FFFFFF&
         Height          =   1066
         Left            =   2866
         MaskColor       =   &H00FFFFFF&
         Picture         =   "1-index.frx":35026
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5148
         Width           =   2470
      End
      Begin VB.CommandButton studentloginbtn 
         BackColor       =   &H00FFFFFF&
         Height          =   1040
         Left            =   2223
         MaskColor       =   &H00FFFFFF&
         Picture         =   "1-index.frx":35A88
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2691
         Width           =   3874
      End
      Begin VB.CommandButton managementl_login 
         BackColor       =   &H00FFFFFF&
         Height          =   1040
         Left            =   2223
         MaskColor       =   &H00FFFFFF&
         Picture         =   "1-index.frx":367A4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3861
         Width           =   3874
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "© all rights reserved by alok kumar (ytsrex media)"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   10.87
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   364
         Left            =   1112
         TabIndex        =   4
         Top             =   6318
         Width           =   5863
      End
      Begin VB.Shape Shape1 
         Height          =   13
         Left            =   2223
         Top             =   1872
         Width           =   3289
      End
      Begin VB.Image Image1 
         Height          =   559
         Left            =   2808
         Picture         =   "1-index.frx":376CA
         Stretch         =   -1  'True
         Top             =   1989
         Width           =   546
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN AREA"
         DragIcon        =   "1-index.frx":3FA20
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   364
         Left            =   2925
         TabIndex        =   1
         Top             =   2106
         Width           =   2587
      End
   End
End
Attribute VB_Name = "Form2_loginarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command5_Click()

End Sub



Private Sub Command2_Click()
contactfrm.Show

End Sub

Private Sub Contactbtn_Click()
contactfrm.Show
Form2_loginarea.Hide

End Sub

Private Sub exit_Click()
End


End Sub

Private Sub helpbtn_Click()
helpDialog.Show
Form2_loginarea.Hide


End Sub

Private Sub managementl_login_Click()
admintfrmLogin.Show
Form2_loginarea.Hide

End Sub

Private Sub studentloginbtn_Click()
studentfrmLogin.Show
Form2_loginarea.Hide



End Sub
