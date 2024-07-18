VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormMasterKaryawan 
   BorderStyle     =   0  'None
   Caption         =   "Form Karyawan"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   Picture         =   "FormMasterKaryawan.frx":0000
   ScaleHeight     =   7785
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
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
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   4800
      Width           =   2415
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
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3230
      Width           =   2295
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
      Height          =   240
      Left            =   2160
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2040
      Width           =   5175
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
      Height          =   240
      Left            =   5040
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   3240
      Width           =   2415
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
      Left            =   4920
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   4750
      Width           =   2775
   End
   Begin VB.Frame FormMasterKaryawan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1560
      TabIndex        =   1
      Top             =   6000
      Width           =   6615
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Input"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
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
         Top             =   240
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
      Left            =   9840
      Top             =   4080
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
      Left            =   9840
      TabIndex        =   0
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Image Image1 
      Height          =   495
      Left            =   9000
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FormMasterKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Input" Then
    Command1.Caption = "Simpan"
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "Batal"
    Else
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Combo1 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim TambahData
    TambahData = "Insert into TBL_Karyawan values ('" & Text1 & "','" & Text2 & "','" & Left(Combo1, 6) & "','" & Text3 & "','" & Text4 & "')"
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
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Combo1 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim EditData
    EditData = "update TBL_Karyawan set NamaKaryawan = '" & Text2 & "',KodeJabatan = '" & Combo1 & "', AlamatKaryawan = '" & Text3 & "', TelpKaryawan = '" & Text4 & "' where NIK='" & Text1 & "'"
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
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Combo1 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim HapusData As String
    HapusData = "Delete from TBL_Karyawan where NIK='" & Text1 & "'"
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
    Else
    Call KondisiAwal
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
Text4 = ""
Command1.Caption = "Input"
Command2.Caption = "Edit"
Command3.Caption = "Hapus"
Command4.Caption = "Tutup"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True

Call BukaDB
RSJabatan.Open "tbl_jabatan", Koneksi
Combo1.Clear
Do Until RSJabatan.EOF
    Combo1.AddItem RSJabatan!KodeJabatan
    RSJabatan.MoveNext
Loop

End Sub

Sub MunculData()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "TBL_Karyawan"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BukaDB
    RSKaryawan.Open "Select * From TBL_Karyawan where NIK= '" & Text1 & "'", Koneksi
    If Not RSKaryawan.EOF Then
    Text2 = RSKaryawan!NamaKaryawan
    Text3 = RSKaryawan!AlamatKaryawan
    Text4 = RSKaryawan!TelpKaryawan
    Combo1 = RSKaryawan!KodeJabatan
     Else
    MsgBox "Data Tidak Ada!"
    End If
End If
End Sub

