VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pinjam 
   Caption         =   "Peminjaman Film"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form3"
   ScaleHeight     =   5760
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   4920
      TabIndex        =   3
      Top             =   3120
      Width           =   1250
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   6360
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   1800
      TabIndex        =   6
      Top             =   2760
      Width           =   750
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   750
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   750
   End
   Begin VB.TextBox TxtNomorAgt 
      Height          =   350
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1100
   End
   Begin MSDataGridLib.DataGrid DG1 
      Bindings        =   "Pinjam.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "Nomor"
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
         DataField       =   "Kode"
         Caption         =   "Kode"
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
      BeginProperty Column02 
         DataField       =   "Judul"
         Caption         =   "Judul"
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
      BeginProperty Column03 
         DataField       =   "Jumlah"
         Caption         =   "Jumlah"
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
      BeginProperty Column04 
         DataField       =   "Tarif"
         Caption         =   "Tarif"
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
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   375
      Left            =   120
      Top             =   3240
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Transaksi"
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
   Begin MSDataGridLib.DataGrid DG2 
      Bindings        =   "Pinjam.frx":0011
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "NomorPjm"
         Caption         =   "NomorPjm"
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
         DataField       =   "NomorFlm"
         Caption         =   "Nomor Flm"
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
      BeginProperty Column02 
         DataField       =   "Judul"
         Caption         =   "Judul Film"
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
      BeginProperty Column03 
         DataField       =   "JumlahFlm"
         Caption         =   "Jumlah"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3495,118
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DTCari 
      Height          =   375
      Left            =   2160
      Top             =   3240
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DTCari"
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
   Begin VB.Label LblKembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4920
      TabIndex        =   22
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   4080
      TabIndex        =   21
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   4080
      TabIndex        =   20
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label LblTotalHrg 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4920
      TabIndex        =   19
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   4080
      TabIndex        =   18
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telah Pinjam"
      Height          =   345
      Left            =   2280
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LbltelahPjm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3360
      TabIndex        =   16
      Top             =   120
      Width           =   540
   End
   Begin VB.Label LblNamaAgt 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2280
      TabIndex        =   14
      Top             =   480
      Width           =   3900
   End
   Begin VB.Label LblTotalPjm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3360
      TabIndex        =   13
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah"
      Height          =   345
      Left            =   2640
      TabIndex        =   12
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   11
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label LblNomorPjm 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1200
      TabIndex        =   10
      Top             =   120
      Width           =   1100
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Anggota"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Pinjam"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "Pinjam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
On Error Resume Next
    DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOVCD.mdb"
    DT.RecordSource = "Transaksi"
    Set DG1.DataSource = DT
    DG1.Refresh

    Call BukaDB
    RSFilm.Open "SELECT * FROM Film WHERE STOK>0", Conn
    List1.Clear
    Do Until RSFilm.EOF
        List1.AddItem RSFilm!Judul & Space(50) & RSFilm!Nomorflm
        RSFilm.MoveNext
    Loop
    
    Call AutoNomor
    LblTanggal.Caption = Format(Date, "dd-mm-yyyy")
    Call Tabel_Kosong
    DT.Recordset.MoveFirst
    DG1.Col = 1
    LblTotalHrg = 0
    TxtDibayar = 0
    LblKembali = 0
End Sub

Private Sub Form_Load()
Call BukaDB
End Sub

Function Tabel_Kosong()
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    DT.Recordset.Delete
    DT.Recordset.MoveNext
Loop
For i = 1 To 1
    DT.Recordset.AddNew
    DT.Recordset!Nomor = i
    DT.Recordset.Update
Next i
End Function

Private Sub AutoNomor()
Call BukaDB
RSPinjam.Open "select * from Pinjam Where NomorPjm In(Select Max(NomorPjm)From Pinjam)Order By NomorPjm Desc", Conn
RSPinjam.Requery
    Dim Urutan As String * 8
    Dim Hitung As Long
    With RSPinjam
        If .EOF Then
            Urutan = Format(Date, "yymmdd") + "01"
            LblNomorPjm = Urutan
        Else
            If Left(!NomorPjm, 6) <> Format(Date, "yymmdd") Then
                Urutan = Format(Date, "yymmdd") + "01"
            Else
                Hitung = (!NomorPjm) + 1
                Urutan = Format(Date, "yymmdd") + Right("00" & Hitung, 2)
            End If
        End If
        LblNomorPjm = Urutan
    End With
End Sub

'Private Sub LblTotalHrg_Change()
'If LblTotalHrg = "" Or LblTotalHrg = 0 Then
'    CmdSimpan.Enabled = True
'Else
'    CmdSimpan.Enabled = False
'End If
'End Sub

'Private Sub AutoNomor()
'Call BukaDB
'RSPinjam.Open "select * from Pinjam Where NomorPjm In(Select Max(NomorPjm)From Pinjam)Order By NomorPjm Desc", Conn
'RSPinjam.Requery
'    Dim Urutan As String * 8
'    Dim Hitung As Long
'    With RSPinjam
'        If .EOF Then
'            Urutan = Format(Date, "yymmdd") + "01"
'            LblNomorPjm = Urutan
'        Else
'            If Left(!NomorPjm, 6) <> Format(Date, "yymmdd") + "01" Then
'                Urutan = Format(Date, "yymmdd") + "01"
'            Else
'                Hitung = (!NomorPjm) + 1
'                Urutan = Format(Date, "yymmdd") + Right("00" & Hitung, 2)
'            End If
'        End If
'        LblNomorPjm = Urutan
'    End With
'End Sub

Private Sub TxtNomorAgt_KeyPress(Keyascii As Integer)
TxtNomorAgt.MaxLength = 4
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Call BukaDB
    RSAnggota.Open "Select * from anggota where nomoragt='" & TxtNomorAgt & "'", Conn

    If Not RSAnggota.EOF Then
        LblNamaAgt.Caption = RSAnggota!Namaagt
        DG1.SetFocus
        DG1.Col = 1
    Else
        MsgBox "Nomor anggota tidak terdaftar"
        TxtNomorAgt.SetFocus
        Exit Sub
    End If
        
    DTCari.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOVCD.mdb"
    DTCari.RecordSource = "select Film.Judul,detailpjm.JumlahFlm from Film,detailpjm,anggota where Film.NomorFlm=detailpjm.NomorFlm and nomoragt=' " & TxtNomorAgt & "'"
    DTCari.Refresh
    DG2.Refresh
    LbltelahPjm.Caption = DTCari.Recordset.RecordCount
    
    Call TelahPjm
    
    If TelahPjm = 0 Or LbltelahPjm = "" Then
        DG1.SetFocus
        DG1.Col = 1
    Else
        Call Pinjaman
        DG1.SetFocus
        DG1.Col = 1
        DG2.Visible = True
        Exit Sub
    End If
End If
End Sub

Function TelahPjm()
    On Error Resume Next
    Set TTLPjm = New ADODB.Recordset
    TTLPjm.Open "SELECT sum(TOTALPJM) AS JUMTOTAL FROM PINJAM WHERE NOMORAGT='" & TxtNomorAgt & "'", Conn
    TelahPjm = TTLPjm!JumTotal
    LbltelahPjm.Caption = TelahPjm
End Function

Sub Pinjaman()
    DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOVCD.mdb"
    DTCari.RecordSource = "Select Distinct Detailpjm.Nomorpjm,Film.Nomorflm,Judul,Jumlahflm From Anggota,Pinjam,Film,Detailpjm Where Film.Nomorflm=Detailpjm.Nomorflm And Pinjam.Nomorpjm=Left(Detailpjm.Nomorpjm,8) And Anggota.Nomoragt=Pinjam.Nomoragt And Anggota.Nomoragt='" & TxtNomorAgt & "'"
    DTCari.Refresh
    LbltelahPjm.Caption = DTCari.Recordset.RecordCount
End Sub

Private Sub DG1_AfterColEdit(ByVal ColIndex As Integer)
If DG1.Col = 1 Then
    Call BukaDB
    RSFilm.Open "Select * from Film where NomorFlm='" & DT.Recordset!Kode & "'", Conn
    If RSFilm.EOF Then
        Pesan = MsgBox("Kode Flm Tidak Terdaftar")
        DG1.Col = 1
        Exit Sub
    End If
    DT.Recordset!Kode = RSFilm!Nomorflm
    DT.Recordset!Judul = RSFilm!Judul
    DT.Recordset!Jumlah = 1
    DT.Recordset!tarif = RSFilm!tarif
    Call Tambah_Baris
    DT.Recordset.MoveNext
    DG1.Col = 1
    DT.Recordset.MoveLast
    Call JumlahHarga
    Call JumlahItem
End If

If DG1.Col = 3 Then
    DT.Recordset!Jumlah = DT.Recordset!Jumlah
    DT.Recordset.Update
    DT.Recordset.MoveNext
    DG1.Col = 1
    LblTotalPjm.Caption = Format(TotalPjm, "###")
    Call JumlahHarga
    Call JumlahItem
    
End If

End Sub

Function Tambah_Baris()
For i = DT.Recordset.RecordCount To DT.Recordset.RecordCount
    DT.Recordset.AddNew
    DT.Recordset!Nomor = i + 1
    DT.Recordset.Update
Next i
End Function

Private Sub TxtDibayar_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If TxtDibayar = "" Or Val(TxtDibayar) < (LblTotalHrg) Then
            MsgBox "Jumlah Pembayaran Kurang"
            TxtDibayar.SetFocus
        Else
            TxtDibayar = Format(TxtDibayar, "###,###,###")
            If TxtDibayar = LblTotalHrg Then
                LblKembali = TxtDibayar - LblTotalHrg
            Else
                LblKembali = Format(TxtDibayar - LblTotalHrg, "###,###,###")
            End If
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub CmdSimpan_Keypress(Keyascii As Integer)
    If Keyascii = 27 Then
        CmdSimpan.Enabled = False
        TxtDibayar = ""
        TxtDibayar.SetFocus
    End If
End Sub

Private Sub cmdSimpan_Click()
If LblTotalPjm.Caption = "" Then
    MsgBox "Tidak ada transaksi peminjaman"
    TxtNomorAgt.SetFocus
    Exit Sub
ElseIf TxtDibayar = "" Then
    MsgBox "pembayaran Kurang"
    TxtDibayar.SetFocus
    Exit Sub
End If

'simpan ke tabel pinjam
Dim SQLInput1 As String
SQLInput1 = "Insert Into Pinjam(NomorPjm,TanggalPjm,TotalPjm,TotalHrg,Dibayar,Kembali,Nomoragt)" & _
"values('" & LblNomorPjm.Caption & "','" & LblTanggal.Caption & "','" & LblTotalPjm.Caption & "','" & LblTotalHrg.Caption & "','" & TxtDibayar & "','" & LblKembali.Caption & "','" & TxtNomorAgt & "')"
Conn.Execute (SQLInput1)

'simpan ke tabel detailpjm
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    If DT.Recordset!Kode <> vbNullString Then
        Dim SQLInput2 As String
        SQLInput2 = "Insert Into DetailPjm(NomorPjm,NomorFlm,JumlahFlm) " & _
        "values ('" & LblNomorPjm.Caption + DT.Recordset!Nomor & "','" & DT.Recordset!Kode & "','" & DT.Recordset!Jumlah & "')"
        Conn.Execute (SQLInput2)
    End If
DT.Recordset.MoveNext
Loop
    
'Pengurangan Jumlah Flm
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    If DT.Recordset!Kode <> vbNullString Then
        Call BukaDB
        RSFilm.Open "Select * from Film where NomorFlm='" & DT.Recordset!Kode & "'", Conn
        If Not RSFilm.EOF Then
            Dim kurangi As String
            kurangi = "update Film set stok='" & RSFilm!Stok - DT.Recordset!Jumlah & "' where NomorFlm='" & DT.Recordset!Kode & "'"
            Conn.Execute (kurangi)
        End If
    End If
DT.Recordset.MoveNext
Loop
Call Bersihkan
Form_Activate
cmdbatal_Click
End Sub

Sub Bersihkan()
TxtNomorAgt = ""
LblNamaAgt.Caption = ""
LblTotalPjm.Caption = ""
LbltelahPjm.Caption = ""
LblTotalHrg.Caption = ""
TxtDibayar = ""
LblKembali.Caption = ""
End Sub

'Function TotalPjm()
'On Error Resume Next
'DT.Recordset.MoveFirst
'Jumlah = 0
'Do While Not DT.Recordset.EOF And DT.Recordset!Jumlah <> 0
'    Jumlah = Jumlah + DT.Recordset!Jumlah
'    DT.Recordset.MoveNext
'    LblTotalPjm = Format(Jumlah, "#,###,###")
'Loop
'End Function
'
'
'Function TotalHrg()
'On Error Resume Next
'DT.Recordset.MoveFirst
'total = 0
'Do While Not DT.Recordset.EOF And DT.Recordset!tarif <> 0
'    total = total + DT.Recordset!tarif
'    DT.Recordset.MoveNext
'    LblTotalHrg = Format(total, "#,###,###")
'Loop
'
'End Function

Private Sub cmdbatal_Click()
On Error Resume Next
Form_Activate
TxtNomorAgt = ""
LblNamaAgt = ""
LblTotalPjm = ""
LbltelahPjm = ""
Call Pinjaman
TxtNomorAgt.SetFocus
End Sub

Private Sub cmdtutup_Click()
Unload Me
End Sub

Private Sub List1_keyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        Call BukaDB
        Dim cari As New ADODB.Recordset
        cari.Open "select * from transaksi where kode='" & Right(List1, 4) & "'", Conn
        cari.Requery
        If Not cari.EOF Then
            Pesan = MsgBox("yakin akan dientri dua kali", vbYesNo)
            If Pesan = vbYes Then
                RSFilm.Open "Select * from Film where nomorflm ='" & Right(List1, 4) & "'", Conn
                RSFilm.Requery
                If Not RSFilm.EOF Then
                    DT.Recordset!Kode = RSFilm!Nomorflm
                    DT.Recordset!Judul = RSFilm!Judul
                    DT.Recordset!Jumlah = 1
                    DT.Recordset!tarif = RSFilm!tarif
                    Call Tambah_Baris
                    DT.Recordset.MoveNext
                    DG1.Col = 1
                    DT.Recordset.MoveLast
                    Call JumlahHarga
                    Call JumlahItem
                End If
            End If
        Else
            Conn.Close
            Call BukaDB
            RSFilm.Open "Select * from Film where nomorflm ='" & Right(List1, 4) & "'", Conn
            RSFilm.Requery
            If Not RSFilm.EOF Then
                DT.Recordset!Kode = RSFilm!Nomorflm
                DT.Recordset!Judul = RSFilm!Judul
                DT.Recordset!Jumlah = 1
                DT.Recordset!tarif = RSFilm!tarif
                Call Tambah_Baris
                DT.Recordset.MoveNext
                DG1.Col = 1
                DT.Recordset.MoveLast
                Call JumlahHarga
                Call JumlahItem
            End If
        End If
    End If
End Sub

Function JumlahHarga()
DT.Recordset.MoveFirst
A = 0
Do While Not DT.Recordset.EOF And DT.Recordset!tarif <> vbNullString
    A = A + DT.Recordset!tarif
    DT.Recordset.MoveNext
    LblTotalHrg = Format(A, "###,###,###")
Loop
End Function

Function JumlahItem()
DT.Recordset.MoveFirst
A = 0
Do While Not DT.Recordset.EOF And DT.Recordset!Jumlah <> vbNullString
    A = A + DT.Recordset!Jumlah
    DT.Recordset.MoveNext
    LblTotalPjm = A
Loop
End Function

