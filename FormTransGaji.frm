VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormTransGaji 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormTransGaji.frx":0000
   ScaleHeight     =   8310
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Hitung "
      Height          =   375
      Left            =   13320
      TabIndex        =   19
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   18
      Top             =   6720
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   12000
      TabIndex        =   17
      Top             =   7680
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormTransGaji.frx":DFEF
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   8400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4800
      Top             =   9600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"FormTransGaji.frx":E004
      OLEDBString     =   $"FormTransGaji.frx":E091
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TBL_GAJI"
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
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   11640
      TabIndex        =   10
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8760
      TabIndex        =   9
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input"
      Height          =   375
      Left            =   13440
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6840
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text8 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   11640
      TabIndex        =   11
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   11640
      TabIndex        =   8
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8760
      TabIndex        =   7
      Top             =   4680
      Width           =   2025
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9720
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8760
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   14280
      Top             =   240
      Width           =   495
   End
   Begin VB.Label LBLJabatan 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label LBLNamaJabatan 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Label LBLNoGaji 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label LBLTanggal 
      BackStyle       =   0  'Transparent
      Caption         =   "LBLTanggal"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label LBLTelp 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   3600
      Width           =   2115
   End
   Begin VB.Label LBLAlamat 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   5760
      Width           =   5775
   End
   Begin VB.Label LBLNama 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   3600
      Width           =   1935
   End
End
Attribute VB_Name = "FormTransGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KoneksiDB As New ADODB.Connection
Dim RSKaryawan As New ADODB.Recordset
Dim RSJabatan As New ADODB.Recordset
Dim RSGaji As New ADODB.Recordset

Private Sub Combo1_Click()
    Call BukaDB
    Set RSKaryawan = New ADODB.Recordset
    RSKaryawan.Open "SELECT * FROM TBL_KARYAWAN WHERE NIK='" & Combo1.Text & "'", KoneksiDB
    If Not RSKaryawan.EOF Then
        LBLNama.Caption = RSKaryawan!NamaKaryawan
        LBLAlamat.Caption = RSKaryawan!AlamatKaryawan
        LBLTelp.Caption = RSKaryawan!TelpKaryawan
        LBLJabatan.Caption = RSKaryawan!KodeJabatan
        
        Set RSJabatan = New ADODB.Recordset
        RSJabatan.Open "SELECT * FROM TBL_JABATAN WHERE KodeJabatan='" & LBLJabatan.Caption & "'", KoneksiDB
        If Not RSJabatan.EOF Then
            Text2.Text = RSJabatan!GajiPokok
            LBLNamaJabatan.Caption = RSJabatan!Jabatan
            
            Text3.Text = "0"
            Text4.Text = "0"
            Text5.Text = "0"
            Text6.Text = "0"
            Text7.Text = "0"
            Text8.Text = "150000"
        End If
        RSJabatan.Close
    End If
    RSKaryawan.Close
End Sub

Private Sub Command1_Click()
    CrystalReport1.ReportFileName = App.Path & "\SlipGaji2.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 0
End Sub

Sub MunculData()
    Call BukaDB
    Adodc1.ConnectionString = KoneksiDB.ConnectionString
    Adodc1.RecordSource = "SELECT * FROM TBL_GAJI"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrorHandler
    
    ' Buka koneksi ke database
    Call BukaDB
    
    ' Periksa apakah semua label dan textbox memiliki nilai yang valid
    If LBLNoGaji.Caption = "" Or LBLTanggal.Caption = "" Or Combo1.Text = "" Or _
       LBLNama.Caption = "" Or LBLJabatan.Caption = "" Or LBLAlamat.Caption = "" Or _
       LBLTelp.Caption = "" Or Text2.Text = "" Or Text3.Text = "" Or _
       Text4.Text = "" Or Text5.Text = "" Or Text7.Text = "" Or _
       Text8.Text = "" Or Text1.Text = "" Then
        MsgBox "Pastikan semua data terisi dengan benar.", vbExclamation
        Exit Sub
    End If
    
    ' Buat perintah SQL untuk menambahkan data
    Dim TambahData1
    TambahData = "INSERT INTO TBL_GAJI (NoSlipGaji, Tanggal, NIK, Nama, KodeJabatan, Jabatan, Alamat, Telp, GajiPokok, JumlahKehadiran, Transport, UangMakan, TunjFungsional, TunjPendidikan, Jamsostek, TotalGaji) " & _
             "VALUES ('" & LBLNoGaji.Caption & "', '" & LBLTanggal.Caption & "', '" & Left(Combo1.Text, 6) & "', '" & LBLNama.Caption & "', '" & LBLJabatan.Caption & "', '" & LBLNamaJabatan.Caption & "', '" & LBLAlamat.Caption & "', '" & LBLTelp.Caption & "', '" & Text2.Text & "', '" & Text3.Text & "', '" & Text4.Text & "', '" & Text5.Text & "', '" & Text6.Text & "', '" & Text7.Text & "', '" & Text8.Text & "', '" & Text1.Text & "')"
    
    ' Eksekusi perintah SQL
    KoneksiDB.Execute TambahData
    
    ' Tampilkan pesan sukses
    MsgBox "Tambah Data Berhasil"
    
    ' Refresh data grid
    Call MunculData

    ' Tutup koneksi database
    KoneksiDB.Close
    Exit Sub

ErrorHandler:
    MsgBox "Terjadi kesalahan: " & Err.Description, vbCritical
    If KoneksiDB.State = adStateOpen Then KoneksiDB.Close
End Sub

Private Sub Command3_Click()
    CrystalReport1.ReportFileName = App.Path & "\SlipGaji2.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 0
End Sub

Private Sub Command4_Click()
  Text9.Text = Val(Text2.Text) + Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) - Val(Text8.Text)
End Sub

Private Sub Form_Load()
    
    LBLTanggal.Caption = Date$
    Call NoOtomatis
    Call MunculNIK
    Call BukaDB
End Sub

Private Sub NoOtomatis()
   Call BukaDB
    Set RSGaji = New ADODB.Recordset
    RSGaji.Open "SELECT * FROM TBL_GAJI WHERE NoSlipGaji IN (SELECT MAX(NoSlipGaji) FROM TBL_GAJI) ORDER BY NoSlipGaji DESC", KoneksiDB
    Dim Urutan As String * 10
    Dim Hitung As Long
    If RSGaji.EOF Then
        Urutan = "GJ00000001"
        LBLNoGaji.Caption = Urutan
    Else
        Hitung = CLng(Right(RSGaji!NoSlipGaji, 8)) + 1
        Urutan = "GJ" & Right("00000000" & Hitung, 8)
        LBLNoGaji.Caption = Urutan
    End If
    RSGaji.Close
End Sub


Private Sub MunculNIK()
    Call BukaDB
    Set RSKaryawan = New ADODB.Recordset
    RSKaryawan.Open "SELECT * FROM TBL_KARYAWAN", KoneksiDB
    Combo1.Clear
    Do Until RSKaryawan.EOF
        Combo1.AddItem RSKaryawan!NIK
        RSKaryawan.MoveNext
    Loop
    RSKaryawan.Close
End Sub

Private Sub Hitung_Click()
    Text1.Text = Val(Text2.Text) + Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) - Val(Text8.Text)
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call BukaDB
        Set RSJabatan = New ADODB.Recordset
        RSJabatan.Open "SELECT * FROM TBL_JABATAN WHERE KodeJabatan = '" & LBLJabatan.Caption & "'", KoneksiDB
        If Not RSJabatan.EOF Then
            Text4.Text = RSJabatan!uangtransport * Val(Text3.Text)
            Text5.Text = "600000"
            Text6.Text = "400000"
            Text7.Text = "500000"
        Else
            MsgBox "Data Tidak Ada!"
        End If
        RSJabatan.Close
    End If
End Sub

Sub BukaDB()
    If KoneksiDB.State = adStateOpen Then KoneksiDB.Close
    KoneksiDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBPenggajian.mdb;"
End Sub


