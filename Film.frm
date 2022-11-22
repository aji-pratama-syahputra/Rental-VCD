VERSION 5.00
Begin VB.Form Film 
   Caption         =   "Data Film"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   LinkTopic       =   "Form2"
   ScaleHeight     =   2310
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTarif 
      Height          =   350
      Left            =   1200
      TabIndex        =   11
      Top             =   840
      Width           =   3200
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   120
      Width           =   1600
   End
   Begin VB.TextBox TxtNomor 
      Height          =   350
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox TxtJudul 
      Height          =   350
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   3200
   End
   Begin VB.TextBox TxtStok 
      Height          =   350
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   3200
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   1000
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1000
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tarif"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Judul"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stok"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1005
   End
End
Attribute VB_Name = "Film"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Form_Activate()
    Call BukaDB
    RSFilm.Open "Film", Conn
    Combo1.Clear
    Do Until RSFilm.EOF
        Combo1.AddItem RSFilm!Nomorflm
        RSFilm.MoveNext
    Loop
End Sub

Sub Form_Load()
    Call BukaDB
    TxtNomor.MaxLength = 4
    TxtJudul.MaxLength = 30
    TxtTarif.MaxLength = 5
    TxtStok.MaxLength = 3
    KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSFilm.Open "Select * From Film where NomorFlm='" & TxtNomor & "'", Conn
End Function

Function CariCombo()
    Call BukaDB
    RSFilm.Open "Select * From Film where NomorFlm='" & Combo1 & "'", Conn
End Function

Private Sub KosongkanText()
    TxtNomor = ""
    TxtJudul = ""
    TxtTarif = ""
    TxtStok = ""
    Combo1.Text = ""
End Sub

Private Sub SiapIsi()
    TxtNomor.Enabled = True
    TxtJudul.Enabled = True
    TxtTarif.Enabled = True
    TxtStok.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    TxtNomor.Enabled = False
    TxtJudul.Enabled = False
    TxtTarif.Enabled = False
    TxtStok.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdHapus.Caption = "&Hapus"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
    CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSFilm
        If Not RSFilm.EOF Then
            TxtJudul = RSFilm!Judul
            TxtStok = RSFilm!Stok
        End If
    End With
End Sub

Private Sub AutoNomor()
Call BukaDB
RSFilm.Open ("select * from Film Where NomorFlm In(Select Max(NomorFlm)From Film)Order By NomorFlm Desc"), Conn
RSFilm.Requery
    Dim Urutan As String * 4
    Dim Hitung As Long
    With RSFilm
        If .EOF Then
            Urutan = "F" + "001"
            TxtNomor = Urutan
        Else
            Hitung = Right(!Nomorflm, 3) + 1
            Urutan = "F" + Right("000" & Hitung, 3)
        End If
        TxtNomor = Urutan
    End With
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Call AutoNomor
        TxtNomor.Enabled = False
        TxtJudul.SetFocus
    Else
        If TxtNomor = "" Or TxtJudul = "" Or TxtTarif = "" Or TxtStok = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Film (NomorFlm,Judul,Tarif,Stok) values ('" & TxtNomor & "','" & TxtJudul & "','" & TxtTarif & "','" & TxtStok & "')"
            Conn.Execute SQLTambah
            Call KondisiAwal
            Call Form_Activate
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If CmdEdit.Caption = "&Edit" Then
        CmdInput.Enabled = False
        CmdEdit.Caption = "&Simpan"
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        'TxtNomor.SetFocus
        TxtNomor.Enabled = False
        Combo1.SetFocus
    Else
        If TxtJudul = "" Or TxtTarif = "" Or TxtStok = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Film Set Judul= '" & TxtJudul & "',tarif= '" & TxtTarif & "', stok='" & TxtStok & "' where NomorFlm='" & TxtNomor & "'"
            Conn.Execute SQLEdit
            Call KondisiAwal
            Call Form_Activate
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If CmdHapus.Caption = "&Hapus" Then
        CmdInput.Enabled = False
        CmdEdit.Enabled = False
        CmdTutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        'TxtNomor.SetFocus
        TxtNomor.Enabled = False
        Combo1.SetFocus
    End If
End Sub

Private Sub cmdtutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub TxtNomor_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(TxtNomor) < 4 Then
        MsgBox "Kode Harus 4 Digit"
        TxtNomor.SetFocus
    Else
        TxtJudul.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
            If Not RSFilm.EOF Then
                TampilkanData
                MsgBox "Kode Film Sudah Ada"
                KosongkanText
                TxtNomor.SetFocus
            Else
                TxtJudul.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSFilm.EOF Then
                TampilkanData
                TxtNomor.Enabled = False
                TxtJudul.SetFocus
            Else
                MsgBox "Kode Film Tidak Ada"
                TxtNomor = ""
                TxtNomor.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSFilm.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Film where NomorFlm= '" & TxtNomor & "'"
                    Conn.Execute SQLHapus
                    KondisiAwal
                    'CmdRefresh.SetFocus
                Else
                    KondisiAwal
                    CmdHapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                TxtNomor.SetFocus
            End If
    End If
End If
End Sub

Private Sub txtjudul_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then TxtTarif.SetFocus
End Sub

Private Sub txttarif_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then TxtStok.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub txtstok_keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Combo1_Click()
Call BukaDB
RSFilm.Open "select * from Film where NomorFlm='" & Combo1 & "'", Conn
If Not RSFilm.EOF Then
    With RSFilm
    If Not RSFilm.EOF Then
        TxtNomor = RSFilm!Nomorflm
        TxtJudul = RSFilm!Judul
        TxtTarif = RSFilm!tarif
        TxtStok = RSFilm!Stok
    End If
    End With
End If
End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
If Keyascii = 13 Then
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
        If Not RSFilm.EOF Then
            TampilkanData
            TxtNomor.Enabled = False
            TxtJudul.SetFocus
        Else
            MsgBox "Nomor Rekening Film Tidak Ada"
            TxtNomor.SetFocus
            Exit Sub
        End If
    
    ElseIf CmdHapus.Caption = "&Hapus" Then
        Call CariData
        If Not RSFilm.EOF Then
            TampilkanData
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                Dim hapus As String
                hapus = "Delete From Film where NomorFlm= '" & TxtNomor & "'"
                Conn.Execute hapus
                KondisiAwal
                Form_Activate
            Else
                KondisiAwal
            End If
        Else
            MsgBox "Data Tidak ditemukan"
            TxtNomor.SetFocus
        End If
    End If
End If
End Sub

