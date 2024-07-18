VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormMenuUtama 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Menu Utama Aplikasi Penggajian"
   ClientHeight    =   9660
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   Picture         =   "FormMenuUtama.frx":0000
   ScaleHeight     =   389.713
   ScaleMode       =   0  'User
   ScaleWidth      =   1264
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   9000
   End
   Begin MSComctlLib.StatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9285
      Width           =   18960
      _ExtentX        =   33443
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "KODE"
            TextSave        =   "KODE"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "NAMA"
            TextSave        =   "NAMA"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "LEVEL"
            TextSave        =   "LEVEL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "JAM"
            TextSave        =   "JAM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "TANGGAL"
            TextSave        =   "TANGGAL"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnFile 
      Caption         =   "File"
      Begin VB.Menu MnLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu MnLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu MnKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu MnMaster 
      Caption         =   "Master"
      Begin VB.Menu MnAdmin 
         Caption         =   "Admin"
      End
      Begin VB.Menu MnJabatan 
         Caption         =   "Jabatan"
      End
      Begin VB.Menu MnKaryawan 
         Caption         =   "Karyawan"
      End
   End
   Begin VB.Menu MnTransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu MnTransGaji 
         Caption         =   "Penggajian"
      End
   End
End
Attribute VB_Name = "FormMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call Terkunci
End Sub

Private Sub MnAdmin_Click()
FormMasterAdmin.Show vbModal
End Sub

Private Sub MnJabatan_Click()
FormMasterJabatan.Show vbModal
End Sub

Private Sub MnKaryawan_Click()
FormMasterKaryawan.Show vbModal
End Sub

Private Sub MnKeluar_Click()
End
End Sub

Sub Terkunci()
MnLogin.Enabled = True
MnMaster.Enabled = False
MnTransaksi.Enabled = False
STBar.Panels(2) = ""
STBar.Panels(4) = ""
STBar.Panels(6) = ""
End Sub

Private Sub MnLaporan_Click()

End Sub

Private Sub MnLogin_Click()
FormLogin.Show vbModal
End Sub

Private Sub MnLogout_Click()
Call Terkunci
End Sub

Private Sub MnTransGaji_Click()
FormTransGaji.Show vbModal
End Sub

Private Sub Timer1_Timer()
STBar.Panels(8) = Time
STBar.Panels(10) = Date
End Sub
