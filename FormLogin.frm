VERSION 5.00
Begin VB.Form FormLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11490
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Palette         =   "FormLogin.frx":0000
   Picture         =   "FormLogin.frx":72B5
   ScaleHeight     =   6465
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   210
      Left            =   6960
      TabIndex        =   1
      Text            =   "Password"
      Top             =   3600
      Width           =   2895
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
      Height          =   240
      Left            =   6960
      TabIndex        =   0
      Text            =   "KodeAdmin"
      Top             =   2655
      Width           =   2895
   End
   Begin VB.Image BTNCANCEL 
      Height          =   495
      Left            =   10800
      Top             =   120
      Width           =   495
   End
   Begin VB.Image BTNLOGIN 
      Height          =   495
      Left            =   7320
      Top             =   4560
      Width           =   2055
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Terbuka()
FormMenuUtama.MnLogin.Enabled = False
FormMenuUtama.MnLogout.Enabled = True
FormMenuUtama.MnMaster.Enabled = True
FormMenuUtama.MnTransaksi.Enabled = True
End Sub

Private Sub BTNCANCEL_Click()
Unload Me
End Sub

Private Sub BTNLOGIN_Click()
Call CariData
End Sub



Private Sub Form_Activate()
Text1 = ""
Text2 = ""
Text1.SetFocus
End Sub

Function CariData()
Call BukaDB


RSAdmin.Open "Select * From TBL_ADMIN where KodeAdmin = '" & Text1 & "' and PasswordAdmin = '" & Text2 & "'", Koneksi
If RSAdmin.EOF Then
    MsgBox "Kode Admin atau Password Salah!"
    Text1.SetFocus
    Else
    Unload Me
    FormMenuUtama.Show
    FormMenuUtama.STBar.Panels(2) = RSAdmin!KodeAdmin
    FormMenuUtama.STBar.Panels(4) = RSAdmin!NamaAdmin
    FormMenuUtama.STBar.Panels(6) = RSAdmin!LevelAdmin
    
    Call Terbuka
End If
End Function

Private Sub Form_Load()
Text2.PasswordChar = "*"
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BTNLOGIN_Click
End If
End Sub
