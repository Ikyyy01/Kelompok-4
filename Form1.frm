VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormMasterKaryawan 
   Caption         =   "Form Karyawan"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   -1440
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   -960
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   -480
      Width           =   4695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   0
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   -480
      TabIndex        =   1
      Top             =   720
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "Input"
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   615
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hapus"
         Height          =   615
         Left            =   3360
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Tutup"
         Height          =   615
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4800
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   840
      TabIndex        =   0
      Top             =   3120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
            LCID            =   2057
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
            LCID            =   2057
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Admin"
      Height          =   495
      Left            =   -120
      TabIndex        =   13
      Top             =   -1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
      Height          =   495
      Left            =   -120
      TabIndex        =   12
      Top             =   -960
      Width           =   2055
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      Height          =   495
      Left            =   -120
      TabIndex        =   11
      Top             =   -480
      Width           =   2055
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level"
      Height          =   375
      Left            =   -120
      TabIndex        =   10
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "FormMasterKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
