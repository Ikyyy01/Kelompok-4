VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormMasterAdmin 
   BorderStyle     =   0  'None
   Caption         =   "Form Admin"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "FormMasterAdmin.frx":0000
   Picture         =   "FormMasterAdmin.frx":2964C
   ScaleHeight     =   7800
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1560
      TabIndex        =   5
      Top             =   6480
      Width           =   6975
      Begin VB.CommandButton Command4 
         Caption         =   "tutup"
         Height          =   495
         Left            =   5640
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Input"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8280
      ScaleHeight     =   615
      ScaleWidth      =   1575
      TabIndex        =   10
      Top             =   6720
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10680
      Top             =   6960
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
      Height          =   5415
      Left            =   11040
      TabIndex        =   4
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12648447
      ForeColor       =   0
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   5640
      Width           =   5775
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   4440
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   3240
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2000
      Width           =   5295
   End
   Begin VB.Image btntutup 
      Height          =   495
      Left            =   9240
      Top             =   120
      Width           =   495
   End
   Begin VB.Image btnhapus 
      Height          =   375
      Left            =   6120
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Image btnedit 
      Height          =   375
      Left            =   4080
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image btnsimpan 
      Height          =   375
      Left            =   1680
      Top             =   6720
      Width           =   1335
   End
End
Attribute VB_Name = "FormMasterAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btntutup_Click()
Unload Me
End Sub


Private Sub Command1_Click()
If Command1.Caption = "Input" Then
    Command1.Caption = "Simpan"
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "Batal"
    Else
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim TambahData
    TambahData = "Insert into TBL_ADMIN values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "')"
    Koneksi.Execute TambahData
    MsgBox "Tambah Data Berhasil"
    Call KondisiAwal
    Call MunculData
    End If
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Edit" Then
    Command2.Caption = "Simpan"
    Command1.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "Batal"
    Else
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim EditData
    EditData = "update TBL_ADMIN set NamaAdmin = '" & Text2 & "',PasswordAdmin = '" & Text3 & "',LevelAdmin = '" & Combo1 & "' where KodeAdmin='" & Text1 & "'"
    Koneksi.Execute EditData
    MsgBox "Update Data Berhasil"
    Call KondisiAwal
    Call MunculData
    End If
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Hapus" Then
    Command3.Caption = "Delete"
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Caption = "Batal"
    Else
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim HapusData As String
    HapusData = "Delete from TBL_ADMIN where KodeAdmin='" & Text1 & "'"
    Koneksi.Execute HapusData
    MsgBox "Hapus Data Berhasil"
    Call KondisiAwal
    Call MunculData
    End If
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Tutup" Then
    Me.Hide
    End
End If
End Sub

 Private Sub Form_Load()
Call KondisiAwal
Call MunculData
End Sub

Sub KondisiAwal()
Text1 = ""
Text2 = ""
Text3 = ""
Combo1.Clear
Combo1.AddItem "ADMIN"
Combo1.AddItem "USER"
Text3.PasswordChar = "*"
Command1.Caption = "Input"
Command2.Caption = "Edit"
Command3.Caption = "Hapus"
Command4.Caption = "Tutup"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Sub MunculData()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "TBL_ADMIN"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BukaDB
    RSAdmin.Open "Select * From TBL_Admin where KodeAdmin = '" & Text1 & "'", Koneksi
    If Not RSAdmin.EOF Then
    Text2 = RSAdmin!NamaAdmin
    Text3 = RSAdmin!PasswordAdmin
    Combo1 = RSAdmin!LevelAdmin
     Else
    MsgBox "Data Tidak Ada!"
    End If
End If
End Sub



