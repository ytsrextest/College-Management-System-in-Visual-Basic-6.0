VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AdminMainForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Management Dashboard"
   ClientHeight    =   6604
   ClientLeft      =   195
   ClientTop       =   793
   ClientWidth     =   11557
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AdminMainForm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "AdminMainForm.frx":1084A
   ScaleHeight     =   9178
   ScaleWidth      =   17550
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton resetbtn 
      BackColor       =   &H8000000D&
      Caption         =   "Reset All"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   7254
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7020
      Width           =   1651
   End
   Begin MSDataGridLib.DataGrid resultDataGrid 
      Bindings        =   "AdminMainForm.frx":25226
      Height          =   3991
      Left            =   1170
      TabIndex        =   28
      Top             =   2223
      Width           =   14989
      _ExtentX        =   27628
      _ExtentY        =   7356
      _Version        =   393216
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   22
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Student Result DB"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid registerDataGrid 
      Bindings        =   "AdminMainForm.frx":25246
      Height          =   3991
      Left            =   1170
      TabIndex        =   26
      Top             =   2223
      Width           =   14989
      _ExtentX        =   27628
      _ExtentY        =   7356
      _Version        =   393216
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   22
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Register Student Request"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid routineDataGrid 
      Bindings        =   "AdminMainForm.frx":25260
      Height          =   3991
      Left            =   1170
      TabIndex        =   24
      Top             =   2223
      Width           =   14872
      _ExtentX        =   27413
      _ExtentY        =   7356
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   22
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Student Routine And Notification Update"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid libraryDataGrid 
      Bindings        =   "AdminMainForm.frx":25279
      Height          =   3991
      Left            =   5265
      TabIndex        =   22
      Top             =   2223
      Width           =   7384
      _ExtentX        =   13611
      _ExtentY        =   7356
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   22
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Check Library Pass DB"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc editadminado 
      Height          =   299
      Left            =   3159
      Top             =   1287
      Visible         =   0   'False
      Width           =   1092
      _ExtentX        =   2013
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
      Connect         =   $"AdminMainForm.frx":25292
      OLEDBString     =   $"AdminMainForm.frx":2531E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from admin_data"
      Caption         =   "madobc"
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
   Begin MSDataGridLib.DataGrid feetransDataGrid 
      Bindings        =   "AdminMainForm.frx":253AA
      Height          =   3991
      Left            =   3510
      TabIndex        =   20
      Top             =   2223
      Width           =   9490
      _ExtentX        =   17492
      _ExtentY        =   7356
      _Version        =   393216
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   22
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "New Fee Payment"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid projectDataGrid 
      Bindings        =   "AdminMainForm.frx":253C4
      Height          =   3991
      Left            =   5616
      TabIndex        =   19
      Top             =   2223
      Width           =   6097
      _ExtentX        =   11238
      _ExtentY        =   7356
      _Version        =   393216
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   22
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Student Project"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid doubtDataGrid 
      Bindings        =   "AdminMainForm.frx":253DD
      Height          =   3991
      Left            =   3276
      TabIndex        =   18
      Top             =   2223
      Width           =   9724
      _ExtentX        =   17924
      _ExtentY        =   7356
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   18
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Student Doubt Class Request"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid studentdbDataGrid 
      Bindings        =   "AdminMainForm.frx":253F4
      Height          =   3991
      Left            =   1170
      TabIndex        =   17
      Top             =   2223
      Width           =   14989
      _ExtentX        =   27628
      _ExtentY        =   7356
      _Version        =   393216
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   22
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Student Database"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc studentdbado 
      Height          =   299
      Left            =   234
      Top             =   1287
      Visible         =   0   'False
      Width           =   1183
      _ExtentX        =   2181
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
      Connect         =   $"AdminMainForm.frx":2540F
      OLEDBString     =   $"AdminMainForm.frx":2549A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from user_data"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid facultydbDataGrid 
      Bindings        =   "AdminMainForm.frx":25525
      Height          =   3874
      Left            =   3393
      TabIndex        =   16
      Top             =   2223
      Width           =   9607
      _ExtentX        =   17708
      _ExtentY        =   7141
      _Version        =   393216
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   18
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Faculty Database"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc facultydbado 
      Height          =   364
      Left            =   1638
      Top             =   1287
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
      Connect         =   $"AdminMainForm.frx":25540
      OLEDBString     =   $"AdminMainForm.frx":255CB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from faculty_details"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton closemanagementdb 
      BackColor       =   &H8000000D&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   9126
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7020
      Width           =   1885
   End
   Begin MSDataGridLib.DataGrid managementDataGrid 
      Bindings        =   "AdminMainForm.frx":25656
      Height          =   3991
      Left            =   1170
      TabIndex        =   14
      Top             =   2223
      Width           =   14976
      _ExtentX        =   27605
      _ExtentY        =   7356
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483645
      ForeColor       =   255
      HeadLines       =   2
      RowHeight       =   18
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Management Database"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "Search Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1534
      Left            =   15561
      TabIndex        =   5
      Top             =   234
      Width           =   1885
      Begin VB.CommandButton seaechcmd 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   364
         Left            =   585
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   702
         Width           =   832
      End
      Begin VB.TextBox searchtxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   364
         Left            =   117
         TabIndex        =   12
         Top             =   234
         Width           =   1651
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Search By Roll. No./ Pin/Name"
         Height          =   364
         Left            =   117
         TabIndex        =   31
         Top             =   1053
         Width           =   1885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Student Request"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1534
      Left            =   7839
      TabIndex        =   4
      Top             =   234
      Width           =   7618
      Begin MSAdodcLib.Adodc feestatusado 
         Height          =   299
         Left            =   6435
         Top             =   1170
         Visible         =   0   'False
         Width           =   1092
         _ExtentX        =   2013
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
         Connect         =   $"AdminMainForm.frx":25671
         OLEDBString     =   $"AdminMainForm.frx":256FC
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from feestatus"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H008080FF&
         Caption         =   "Fee + Fine Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   6435
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   468
         Width           =   1066
      End
      Begin MSAdodcLib.Adodc registerado 
         Height          =   299
         Left            =   5031
         Top             =   1170
         Visible         =   0   'False
         Width           =   1092
         _ExtentX        =   2013
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
         Connect         =   $"AdminMainForm.frx":25787
         OLEDBString     =   $"AdminMainForm.frx":25812
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from register_user"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H008080FF&
         Caption         =   "Register User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   5148
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   468
         Width           =   1183
      End
      Begin MSAdodcLib.Adodc routineado 
         Height          =   299
         Left            =   3744
         Top             =   1053
         Visible         =   0   'False
         Width           =   1092
         _ExtentX        =   2013
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
         Connect         =   $"AdminMainForm.frx":2589D
         OLEDBString     =   $"AdminMainForm.frx":25928
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from routin_notification"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H008080FF&
         Caption         =   "Routine + Notification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   3861
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   468
         Width           =   1183
      End
      Begin MSAdodcLib.Adodc feetransado 
         Height          =   364
         Left            =   2574
         Top             =   1053
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
         Connect         =   $"AdminMainForm.frx":259B3
         OLEDBString     =   $"AdminMainForm.frx":25A3E
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from feetrans"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc projectado 
         Height          =   364
         Left            =   1287
         Top             =   1170
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
         Connect         =   $"AdminMainForm.frx":25AC9
         OLEDBString     =   $"AdminMainForm.frx":25B54
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from project"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc doubtado 
         Height          =   299
         Left            =   117
         Top             =   1170
         Visible         =   0   'False
         Width           =   1092
         _ExtentX        =   2013
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
         Connect         =   $"AdminMainForm.frx":25BDF
         OLEDBString     =   $"AdminMainForm.frx":25C6A
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from student_doubt"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "Fee + Fine Trans."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   2574
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   468
         Width           =   1183
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "Check Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   1287
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   468
         Width           =   1183
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Doubt Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   117
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   468
         Width           =   1066
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Select Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1534
      Left            =   117
      TabIndex        =   3
      Top             =   234
      Width           =   7618
      Begin MSAdodcLib.Adodc semesterresultado 
         Height          =   364
         Left            =   6201
         Top             =   1053
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
         Connect         =   $"AdminMainForm.frx":25CF5
         OLEDBString     =   $"AdminMainForm.frx":25D80
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * From sem_result"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0080FFFF&
         Caption         =   "Semester Result"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   6084
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   468
         Width           =   1417
      End
      Begin MSAdodcLib.Adodc libraryado 
         Height          =   299
         Left            =   4563
         Top             =   1053
         Visible         =   0   'False
         Width           =   1092
         _ExtentX        =   2013
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
         Connect         =   $"AdminMainForm.frx":25E0B
         OLEDBString     =   $"AdminMainForm.frx":25E96
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from library_pass"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Library Pass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   4563
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   468
         Width           =   1417
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Management Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   3042
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   468
         Width           =   1417
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Faculty Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   1638
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   468
         Width           =   1300
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Student Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   117
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   468
         Width           =   1417
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10179
      Top             =   3978
   End
   Begin MSAdodcLib.Adodc fetchadmindetailsado 
      Height          =   364
      Left            =   7956
      Top             =   3978
      Visible         =   0   'False
      Width           =   2119
      _ExtentX        =   3906
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
      Connect         =   $"AdminMainForm.frx":25F21
      OLEDBString     =   $"AdminMainForm.frx":25FAD
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from admin_data"
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
   Begin MSDataGridLib.DataGrid feestatusDataGrid 
      Bindings        =   "AdminMainForm.frx":26039
      Height          =   3991
      Left            =   1170
      TabIndex        =   30
      Top             =   2223
      Width           =   14989
      _ExtentX        =   27628
      _ExtentY        =   7356
      _Version        =   393216
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   22
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Student Fee\Fine Status"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   3523
      Left            =   13455
      Picture         =   "AdminMainForm.frx":26054
      Stretch         =   -1  'True
      Top             =   5850
      Width           =   4108
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   2236
      Left            =   0
      Top             =   0
      Width           =   17797
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.23
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1183
      Left            =   5967
      TabIndex        =   2
      Top             =   6201
      Width           =   6331
   End
   Begin VB.Label Label1 
      DataField       =   "namee"
      DataSource      =   "fetchadmindetailsado"
      Height          =   247
      Left            =   6669
      TabIndex        =   1
      Top             =   4095
      Visible         =   0   'False
      Width           =   1066
   End
   Begin VB.Label Username 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   5499
      TabIndex        =   0
      Top             =   4095
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.Image Image1 
      Height          =   3289
      Left            =   0
      Picture         =   "AdminMainForm.frx":4FF01
      Stretch         =   -1  'True
      Top             =   5967
      Width           =   4225
   End
   Begin VB.Menu dbbacku 
      Caption         =   "Database Backup"
      Begin VB.Menu studentdbbackup 
         Caption         =   "Student DB "
      End
      Begin VB.Menu admindbbacku 
         Caption         =   "Admin DB"
      End
   End
   Begin VB.Menu viewprofile 
      Caption         =   "View Profile"
   End
   Begin VB.Menu changeprofiledetails 
      Caption         =   "Update Profile Details"
   End
   Begin VB.Menu sendmail 
      Caption         =   "Send Mail "
   End
   Begin VB.Menu contactdeveloper 
      Caption         =   "Contact Developer"
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "AdminMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub admindbbacku_Click()
managementdbsaveDialog.Show
End Sub

Private Sub changeprofiledetails_Click()
changeadmindetailsDialog.Show

End Sub

Private Sub closemanagementdb_Click()
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False


closemanagementdb.Visible = False

End Sub

Private Sub Command1_Click()
studentdbDataGrid.Visible = False
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
doubtDataGrid.Visible = True
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
closemanagementdb.Visible = True
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False

End Sub

Private Sub Command10_Click()
registerDataGrid.Visible = True
closemanagementdb.Visible = True
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False

End Sub

Private Sub Command11_Click()
resultDataGrid.Visible = True
closemanagementdb.Visible = True
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
feestatusDataGrid.Visible = False


End Sub

Private Sub Command12_Click()
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = True
closemanagementdb.Visible = True

End Sub

Private Sub Command2_Click()
projectDataGrid.Visible = True
closemanagementdb.Visible = True
studentdbDataGrid.Visible = False
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
doubtDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False

End Sub

Private Sub Command3_Click()
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = True
closemanagementdb.Visible = True
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False


End Sub

Private Sub Command4_Click()
studentdbDataGrid.Visible = True
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False

closemanagementdb.Visible = True
End Sub

Private Sub Command5_Click()
managementDataGrid.Visible = False
facultydbDataGrid.Visible = True
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False

closemanagementdb.Visible = True

End Sub

Private Sub Command6_Click()
managementDataGrid.Visible = True
facultydbDataGrid.Visible = False
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False

closemanagementdb.Visible = True


End Sub

Private Sub Command8_Click()
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = True
closemanagementdb.Visible = True
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False

End Sub

Private Sub Command9_Click()
managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
studentdbDataGrid.Visible = False
doubtDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = True
closemanagementdb.Visible = True
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False

End Sub

Private Sub contactdeveloper_Click()
developercontactDialog.Show

End Sub


Private Sub Form_Load()
Username.Caption = admintfrmLogin.admintxtUserName
fetchadmindetailsado.RecordSource = "select * from admin_data where username='" + Username.Caption + "'"
fetchadmindetailsado.Refresh

Label2.Caption = "Welcome" + Label1 + " TO Admin Area"

managementDataGrid.Visible = False
facultydbDataGrid.Visible = False
doubtDataGrid.Visible = False
closemanagementdb.Visible = False
studentdbDataGrid.Visible = False
projectDataGrid.Visible = False
feetransDataGrid.Visible = False
libraryDataGrid.Visible = False
routineDataGrid.Visible = False
registerDataGrid.Visible = False
resultDataGrid.Visible = False
feestatusDataGrid.Visible = False
resetbtn.Visible = False

'Search Quary


End Sub

Private Sub Label2_Change()
Label2.Caption = "Welcome " + Label1 + " in Admin Dashboard" & vbNewLine & " Date: " & Date & ", Time: " & Time
End Sub

Private Sub logout_Click()
End
End Sub





Private Sub managementDataGrid_Click()
If managementDataGrid.Visible = False Then
closemanagementdb.Visible = False

End If

End Sub

Private Sub resetbtn_Click()

If studentdbDataGrid.Visible = True Then

studentdbado.RecordSource = " select * from user_data "
studentdbado.Refresh
studentdbado.Caption = studentdbado.RecordSource
resetbtn.Visible = False

'FacultyDB
ElseIf facultydbDataGrid.Visible = True Then
facultydbado.RecordSource = " select * from faculty_details "
facultydbado.Refresh
facultydbado.Caption = facultydbado.RecordSource
resetbtn.Visible = False


'ManagementDB
ElseIf managementDataGrid.Visible = True Then
editadminado.RecordSource = " select * from admin_data"
editadminado.Refresh
editadminado.Caption = editadminado.RecordSource
resetbtn.Visible = False


'Labrary pass
ElseIf libraryDataGrid.Visible = True Then
libraryado.RecordSource = " select * from library_pass "
libraryado.Refresh
libraryado.Caption = libraryado.RecordSource
resetbtn.Visible = False

'Semester Result
ElseIf resultDataGrid.Visible = True Then
semesterresultado.RecordSource = " select * from sem_result"
semesterresultado.Refresh
semesterresultado.Caption = semesterresultado.RecordSource
resetbtn.Visible = False

'Doubt Class
ElseIf doubtDataGrid.Visible = True Then
doubtado.RecordSource = " select * from student_doubt"
doubtado.Refresh
doubtado.Caption = doubtado.RecordSource
resetbtn.Visible = False

'Check Project
ElseIf projectDataGrid.Visible = True Then
projectado.RecordSource = " select * from project "
projectado.Refresh
projectado.Caption = projectado.RecordSource
resetbtn.Visible = False

'Fee/ Fine Status
ElseIf feetransDataGrid.Visible = True Then
feetransado.RecordSource = " select * from feetrans "
feetransado.Refresh
feetransado.Caption = feetransado.RecordSource
resetbtn.Visible = False

'Routine and notification
ElseIf routineDataGrid.Visible = True Then
routineado.RecordSource = " select * from routin_notification "
routineado.Refresh
routineado.Caption = routineado.RecordSource
resetbtn.Visible = False


'Register User
ElseIf registerDataGrid.Visible = True Then
registerado.RecordSource = " select * from register_user "
registerado.Refresh
registerado.Caption = registerado.RecordSource
resetbtn.Visible = False

'Fee/Fine Status
ElseIf feestatusDataGrid.Visible = True Then
feestatusado.RecordSource = " select * from feestatus"
feestatusado.Refresh
feestatusado.Caption = feestatusado.RecordSource
resetbtn.Visible = False

'Main Else
Else
MsgBox "Please Select anything then Seaech", vbCritical, "Search Record!"
searchtxt.Text = Empty

End If


End Sub

Private Sub seaechcmd_Click()
'UserDb
If studentdbDataGrid.Visible = True Then

studentdbado.RecordSource = " select * from user_data where rollno='" + searchtxt.Text + "' or namee='" + searchtxt.Text + "'"
studentdbado.Refresh

If studentdbado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty

Else
studentdbado.Caption = studentdbado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'FacultyDB
ElseIf facultydbDataGrid.Visible = True Then
facultydbado.RecordSource = " select * from faculty_details where fname='" + searchtxt.Text + "'"
facultydbado.Refresh

If facultydbado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
facultydbado.Caption = facultydbado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'ManagementDB
ElseIf managementDataGrid.Visible = True Then
editadminado.RecordSource = " select * from admin_data where pin='" + searchtxt.Text + "' or namee='" + searchtxt.Text + "'"
editadminado.Refresh

If editadminado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
editadminado.Caption = editadminado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Labrary pass
ElseIf libraryDataGrid.Visible = True Then
libraryado.RecordSource = " select * from library_pass where rollno='" + searchtxt.Text + "'"
libraryado.Refresh

If libraryado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
libraryado.Caption = libraryado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Semester Result
ElseIf resultDataGrid.Visible = True Then
semesterresultado.RecordSource = " select * from sem_result where rollno='" + searchtxt.Text + "' or namee='" + searchtxt.Text + "'"
semesterresultado.Refresh

If semesterresultado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
semesterresultado.Caption = semesterresultado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Doubt Class
ElseIf doubtDataGrid.Visible = True Then
doubtado.RecordSource = " select * from student_doubt where rollno='" + searchtxt.Text + "'"
doubtado.Refresh

If doubtado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
doubtado.Caption = doubtado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Check Project
ElseIf projectDataGrid.Visible = True Then
projectado.RecordSource = " select * from project where rollno='" + searchtxt.Text + "'"
projectado.Refresh

If projectado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
projectado.Caption = projectado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Fee/ Fine Status
ElseIf feetransDataGrid.Visible = True Then
feetransado.RecordSource = " select * from feetrans where rollno='" + searchtxt.Text + "'"
feetransado.Refresh

If feetransado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
feetransado.Caption = feetransado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Routine and notification
ElseIf routineDataGrid.Visible = True Then
routineado.RecordSource = " select * from routin_notification where course='" + searchtxt.Text + "'"
routineado.Refresh

If routineado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
routineado.Caption = routineado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Register User
ElseIf registerDataGrid.Visible = True Then
registerado.RecordSource = " select * from register_user where rollno='" + searchtxt.Text + "' or namee='" + searchtxt.Text + "'"
registerado.Refresh

If registerado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
registerado.Caption = registerado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Fee/Fine Status
ElseIf feestatusDataGrid.Visible = True Then
feestatusado.RecordSource = " select * from feestatus where rollno='" + searchtxt.Text + "' or sname='" + searchtxt.Text + "'"
feestatusado.Refresh

If feestatusado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
feestatusado.Caption = feestatusado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'Main Else
Else
MsgBox "Please Select anything then Seaech", vbCritical, "Search Record!"
searchtxt.Text = Empty


End If
End Sub

Private Sub sendmail_Click()
adminsendmailDialog.Show

End Sub

Private Sub studentdbbackup_Click()
studentDBbackupDialog.Show

End Sub

Private Sub Timer1_Timer()
Label2.Caption = Date & Time
End Sub

Private Sub Timer2_Timer()
If managementDataGrid.Visible = False Then
closemanagementdb.Visible = False

End If

End Sub

Private Sub Username_Change()
Username.Caption = admintfrmLogin.admintxtUserName
End Sub


Private Sub viewprofile_Click()
ManagementrofileForm.Show

End Sub
