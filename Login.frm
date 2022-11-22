VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3720
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKodeKsr 
      Height          =   350
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   2000
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.TextBox TxtPasswordKsr 
         Height          =   350
         Left            =   1200
         TabIndex        =   1
         Text            =   "LUPA"
         Top             =   720
         Width           =   2000
      End
      Begin VB.TextBox TxtNamaKsr 
         Height          =   350
         Left            =   1200
         TabIndex        =   0
         Text            =   "UUS"
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1005
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A As Byte
Dim B As Byte

Private Sub Form_Load()
'batasi jumlah karakter
TxtNamaKsr.MaxLength = 30
TxtPasswordKsr.MaxLength = 10
'nama dan password diubah menjadi karakter X
'TxtNamaKsr.PasswordChar = "X"
TxtPasswordKsr.PasswordChar = "X"
TxtPasswordKsr.Enabled = False
TxtKodeKsr.Enabled = False
End Sub

Private Sub TxtNamaKsr_KeyPress(Keyascii As Integer)
'ubah karakter jadi besar semua
Keyascii = Asc(UCase(Chr(Keyascii)))
'jika menekan ESC form ditutup
If Keyascii = 27 Then Unload Me
'jika menekan enter setelah mengisi nama, maka..
If Keyascii = 13 Then
    'buka database
    Call BukaDB
    'cari nama kasir yang diketik
    RSkasir.Open "Select NamaKsr from Kasir where NamaKsr ='" & TxtNamaKsr & "'", Conn
    'jika tidak ditemukan, maka
    If RSkasir.EOF Then
        'batasi akses ke nama kasir 3 kali kesempatan
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            'End
            Unload Me
        End If
    Else
        'jika nama kasir benar, maka nama kasir menjadi false
        TxtNamaKsr.Enabled = False
        'password kasir menjadi true dan menjadi fokus kursor
        TxtPasswordKsr.Enabled = True
        TxtPasswordKsr.SetFocus
    End If
End If
End Sub

'coding ini sama dengan nama kasir
Private Sub txtpasswordksr_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
Dim KodeKasir As String
Dim NamaKasir As String
If Keyascii = 13 Then
    Call BukaDB
    RSkasir.Open "Select * from Kasir where NamaKsr ='" & TxtNamaKsr & "' and PasswordKsr='" & TxtPasswordKsr & "'", Conn
    If RSkasir.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            'End
            Unload Me
        End If
    Else
        'jika nama dan password benar, maka...tutup form login
        Unload Me
        'panggil menu utama
        Menu.Show
    End If
End If
End Sub

'program ini berfungsi untuk mengubah
'format tanggal mejadi DDMMYY karena nomor faktur pembelian
'akan muncul otomatis diambil dari tanggal sistem komputer

Sub PeriksaTanggal()
Dim CekTanggal As String
Ulangi:
CekTanggal = Date
If CekTanggal <> Format(Date, "dd/mm/yy") Then
    If MsgBox("Ubah Format tanggal jadi dd/mm/yy di Control Panel, Regional Settings " & vbCrLf & _
    "Customize.., Date, Short Date Style, karena program tidak dapat dijalankan!", vbCritical + vbOKCancel, "Cek Tanggal") = vbOK And CekTanggal <> Format(Date, "dd/mm/yy") Then
        Call Shell("rundll32.exe shell32.dll," & "Control_RunDLL INTL.CPL,,4", 1)
    Else
    End
End If
Pesan = MsgBox("Format Tanggal Sudah diganti..?", vbYesNo, "Konfirmasi")
    If Pesan = vbNo Then
        If CekTanggal <> Format(Date, "dd/mm/yy") Then GoTo Ulangi
        Else
            GoTo Ulangi
        End If
    End If
End Sub

Private Sub Timer1_Timer()
If CekTanggal <> Format(Date, "dd/mm/yy") Then
    PeriksaTanggal
Else
    Exit Sub
End If
End Sub

