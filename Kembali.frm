VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Kembali 
   Caption         =   "Pengembalian Film"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   6960
      TabIndex        =   3
      Top             =   3120
      Width           =   1100
   End
   Begin VB.TextBox TxtNomorAgt 
      Height          =   350
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1500
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   750
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   840
      TabIndex        =   5
      Top             =   2760
      Width           =   750
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   750
   End
   Begin MSDataGridLib.DataGrid DG1 
      Bindings        =   "Kembali.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   8000
      _ExtentX        =   14102
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
      ColumnCount     =   7
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
         DataField       =   "NomorPjm"
         Caption         =   "No Pjm"
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
         DataField       =   "NomorFlm"
         Caption         =   "No Film"
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
      BeginProperty Column04 
         DataField       =   "Tanggal"
         Caption         =   "Tanggal"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "Denda"
         Caption         =   "Denda"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
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
      Bindings        =   "Kembali.frx":0011
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   3201
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "NomorPjm"
         Caption         =   "No Pjm"
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
         DataField       =   "No Film"
         Caption         =   "No Film"
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
         DataField       =   "Tgl Pjm"
         Caption         =   "Tgl Pjm"
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
         DataField       =   "Hrs Kbl"
         Caption         =   "Hrs Kbl"
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
      BeginProperty Column05 
         DataField       =   "Jml Flm"
         Caption         =   "Jml Flm"
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
      BeginProperty Column06 
         DataField       =   "Lama Pjm"
         Caption         =   "Lama Pjm"
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
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DTCari 
      Height          =   375
      Left            =   2160
      Top             =   3240
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Left            =   6960
      TabIndex        =   21
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   6120
      TabIndex        =   20
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   6120
      TabIndex        =   19
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label LblDenda 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6960
      TabIndex        =   18
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Denda"
      Height          =   345
      Left            =   6120
      TabIndex        =   17
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Kembali"
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   5520
      TabIndex        =   15
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Anggota"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label LblNomorKbl 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1200
      TabIndex        =   13
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label LblTanggalKbl 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6600
      TabIndex        =   12
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   2760
      TabIndex        =   11
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label LblTotalKbl 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3480
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label LblNamaAgt 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   5340
   End
   Begin VB.Label LbltelahPjm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telah Pinjam"
      Height          =   345
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "Kembali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOVCD.mdb"
    DT.RecordSource = "Transaksi1"
    Set DG1.DataSource = DT
    DG1.Refresh
    Call AutoNomor
    'LblTanggalKbl.Caption = Format(Date, "dd-mm-yyyy")
    LblTanggalKbl.Caption = Date
    Call Tabel_Kosong
    DT.Recordset.MoveFirst
    DG1.Col = 1
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
RSKembali.Open "select * from Kembali Where NomorKbl In(Select Max(NomorKbl)From Kembali)Order By NomorKbl Desc", Conn
RSKembali.Requery
    Dim Urutan As String * 8
    Dim Hitung As Long
    With RSKembali
        If .EOF Then
            Urutan = Format(Date, "yymmdd") + "01"
            LblNomorKbl = Urutan
        Else
            If Left(!NomorKbl, 6) <> Format(Date, "yymmdd") Then
                Urutan = Format(Date, "yymmdd") + "01"
            Else
                Hitung = (!NomorKbl) + 1
                Urutan = Format(Date, "yymmdd") + Right("00" & Hitung, 2)
            End If
        End If
        LblNomorKbl = Urutan
    End With
End Sub


Private Sub LblDenda_Change()
If Val(LblDenda) = 0 Then
    TxtDibayar.Enabled = False
    TxtDibayar = 0
    LblKembali = 0
    CmdSimpan.Enabled = True
ElseIf LblDenda = "" Then
    TxtDibayar = ""
    TxtDibayar.Enabled = True
    LblKembali = ""
ElseIf LblDenda > 0 Then
    TxtDibayar = ""
    TxtDibayar.Enabled = True
    CmdSimpan.Enabled = False
End If
End Sub

Private Sub TxtNomorAgt_KeyPress(Keyascii As Integer)
On Error Resume Next
TxtNomorAgt.MaxLength = 4
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    LbltelahPjm = ""
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
    
    Call CariPinjaman
      
    If LbltelahPjm = "" Or LbltelahPjm = 0 Then
        MsgBox "'" & LblNamaAgt & "' tidak punya pinjaman"
        'Me.Height = 4455
        TxtNomorAgt.SetFocus
        Exit Sub
    End If
End If
End Sub

Sub CariPinjaman()
DTCari.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOVCD.mdb"
DTCari.RecordSource = "Select Distinct Detailpjm.Nomorpjm,Film.Nomorflm As [No Film],Judul,Tanggalpjm As [Tgl Pjm], (Tanggalpjm+4) As [Hrs Kbl],Jumlahflm As [Jml Flm], (Date()-Tanggalpjm)+1 As [Lama Pjm] From Anggota,Pinjam,Film,Detailpjm Where Film.Nomorflm=Detailpjm.Nomorflm And Pinjam.Nomorpjm=Left(Detailpjm.Nomorpjm,8) And Anggota.Nomoragt=Pinjam.Nomoragt And Anggota.Nomoragt='" & TxtNomorAgt & "'"
DTCari.Refresh
DG2.Refresh
LbltelahPjm.Caption = DTCari.Recordset.RecordCount
End Sub

Private Sub TxtDibayar_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If TxtDibayar = "" Or Val(TxtDibayar) < (LblDenda) Then
            MsgBox "Jumlah Pembayaran Kurang"
            TxtDibayar.SetFocus
        Else
            'TxtDibayar = Format(TxtDibayar, "###,###,###")
            If TxtDibayar = LblDenda Then
                LblKembali = TxtDibayar - LblDenda
            Else
                LblKembali = Format(TxtDibayar - LblDenda, "###,###,###")
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
If LblTotalKbl.Caption = "" Then
    MsgBox "Tidak ada transaksi pengembalian"
    TxtNomorAgt.SetFocus
    Exit Sub
ElseIf TxtDibayar = "" Or TxtDibayar < LblDenda Then
    MsgBox "PEMBAYARAN KURANG "
    'TxtDibayar.SetFocus
    Exit Sub
End If

'simpan ke tabel kembali
Dim SQLInput1 As String
SQLInput1 = "Insert Into kembali(Nomorkbl,Tanggalkbl,Totalkbl,Nomoragt,denda,Dibayar,Kembali)" & _
"values('" & LblNomorKbl & "','" & LblTanggalKbl & "','" & LblTotalKbl & "','" & TxtNomorAgt & "','" & LblDenda & "','" & TxtDibayar & "','" & LblKembali & "')"
Conn.Execute (SQLInput1)

'simpan ke tabel detailkbl
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    If DT.Recordset!NomorPjm <> vbNullString Then
        Dim SQLInput2 As String
        SQLInput2 = "Insert Into Detailkbl(Nomorkbl,NomorFlm,JumlahFlm) " & _
        "values ('" & LblNomorKbl + DT.Recordset!Nomor & "','" & DT.Recordset!Nomorflm & "','" & DT.Recordset!Jumlah & "')"
        Conn.Execute (SQLInput2)
    End If
DT.Recordset.MoveNext
Loop
    
'penambahan Jumlah Film
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    If DT.Recordset!NomorPjm <> vbNullString Then
        Call BukaDB
        RSFilm.Open "Select * from Film where NomorFlm='" & DT.Recordset!Nomorflm & "'", Conn
        If Not RSFilm.EOF Then
            Dim Tambah As String
            Tambah = "update Film set stok='" & RSFilm!Stok + DT.Recordset!Jumlah & "' where nomorFlm='" & DT.Recordset!Nomorflm & "'"
            Conn.Execute (Tambah)
        End If
    End If
DT.Recordset.MoveNext
Loop

'hapus pinjaman
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    If DT.Recordset!NomorPjm <> vbNullString Then
        Call BukaDB
        RSDetailPjm.Open "Select * from detailpjm where nomorpjm='" & DT.Recordset!NomorPjm & "'", Conn
        If Not RSDetailPjm.EOF Then
            Dim hapus As String
            hapus = "delete from detailpjm where nomorpjm ='" & DT.Recordset!NomorPjm & "'"
            Conn.Execute (hapus)
        End If
    End If
DT.Recordset.MoveNext
Loop

'kurangi pinjaman
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    If DT.Recordset!NomorPjm <> vbNullString Then
        Call BukaDB
        RSPinjam.Open "Select * from pinjam where nomorpjm='" & Left(DT.Recordset!NomorPjm, 8) & "'", Conn
        If Not RSPinjam.EOF Then
            Dim kurangi As String
            kurangi = "update pinjam set totalpjm= '" & RSPinjam!TotalPjm - DT.Recordset!Jumlah & " ' where nomorpjm='" & Left(DT.Recordset!NomorPjm, 8) & "' and nomoragt='" & TxtNomorAgt & "'"
            Conn.Execute (kurangi)
        End If
   End If
DT.Recordset.MoveNext
Loop

Bersihkan
Form_Activate
cmdbatal_Click
End Sub

Sub Bersihkan()
TxtNomorAgt = ""
LblNamaAgt.Caption = ""
LblTotalKbl.Caption = ""
LbltelahPjm.Caption = ""
LblDenda = ""
TxtDibayar = ""
LblKembali = ""
End Sub

Private Sub cmdbatal_Click()
Form_Activate
Call Bersihkan
Call CariPinjaman
LbltelahPjm = ""
TxtNomorAgt.SetFocus
End Sub

Private Sub cmdtutup_Click()
Unload Me
End Sub

Private Sub DG2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
    Call BukaDB
    Dim cari As New ADODB.Recordset
    cari.Open "select * from transaksi1 where nomorpjm= '" & DTCari.Recordset!NomorPjm & "'", Conn
    If Not cari.EOF Then
        MsgBox "data jangan dientri dua kali"
        Exit Sub
    Else
        Call SelectAllVisible
    End If
    
'        Call BukaDB
'        Dim CARI As New ADODB.Recordset
'        CARI.Open "SELECT * FROM TRANSAKSI1 WHERE NOMORPJM='" & DTCari.Recordset!NOMORPJM & "'", Conn
'        If Not CARI.EOF Then
'            MsgBox "DATA JANGAN DIENRTI DUA KALI"
'            Exit Sub
'        Else
'            Call SelectAllVisible
'        End If
End Select
End Sub

Sub SelectAllVisible()
'On Error Resume Next
    DT.Recordset!NomorPjm = DG2.Columns(0)
    DT.Recordset!Nomorflm = DG2.Columns(1)
    DT.Recordset!Judul = DG2.Columns(2)
    DT.Recordset!Tanggal = DG2.Columns(3)
    DT.Recordset!Jumlah = DG2.Columns(5)

    If CDate(DT.Recordset!Tanggal) + 5 > 5 Then
        DT.Recordset!Denda = (CDate(LblTanggalKbl) - (DT.Recordset!Tanggal) - 4) * 500 * DT.Recordset!Jumlah
    End If

    If DT.Recordset!Denda < 0 Then
        DT.Recordset!Denda = 0
    End If

    Call Tambah_Baris
    DT.Recordset.MoveNext
    DG1.Col = 1
    DT.Recordset.MoveLast
    Call TotalKbl
    Call JmlDenda

    'LblTotalKbl = TotalKbl
    'LblDenda = Str(JmlDenda)
End Sub

Function Tambah_Baris()
For i = DT.Recordset.RecordCount To DT.Recordset.RecordCount
    DT.Recordset.AddNew
    DT.Recordset!Nomor = i + 1
    DT.Recordset.Update
Next i
End Function

Function TotalKbl()
DT.Recordset.MoveFirst
A = 0
Do While Not DT.Recordset.EOF And DT.Recordset!Jumlah <> vbNullString
    A = A + DT.Recordset!Jumlah
    DT.Recordset.MoveNext
    LblTotalKbl = A
Loop
End Function

Function JmlDenda()
DT.Recordset.MoveFirst
A = 0
Do While Not DT.Recordset.EOF And DT.Recordset!NomorPjm <> vbNullString
    A = A + DT.Recordset!Denda
    DT.Recordset.MoveNext
    LblDenda = A
Loop
End Function

