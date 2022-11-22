VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Laporan 
   Caption         =   "Laporan Transaksi"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Laporan Pengembalian"
      Height          =   2775
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   2850
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   2280
         Width           =   1250
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   1920
         Width           =   1250
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   1320
         Width           =   1250
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   960
         Width           =   1250
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal Awal"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal Akhir"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1245
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   3120
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laporan Peminjaman"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2850
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   2280
         Width           =   1250
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   1920
         Width           =   1250
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1320
         Width           =   1250
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   960
         Width           =   1250
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal Akhir"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal Awal"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1245
      End
   End
End
Attribute VB_Name = "Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call BukaDB
RSPinjam.Open "Select Distinct TanggalPjm From Pinjam order By 1", Conn
RSPinjam.Requery
Do Until RSPinjam.EOF
    Combo1.AddItem RSPinjam!TanggalPjm
    Combo2.AddItem Format(RSPinjam!TanggalPjm, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSPinjam!TanggalPjm, "YYYY ,MM, DD")
    RSPinjam.MoveNext
Loop

For i = 1 To 12
    Combo4.AddItem i
Next i
For i = 1 To 10
    Combo5.AddItem 2000 + i
Next i

RSKembali.Open "Select Distinct TanggalKbl From Kembali order By 1", Conn
RSKembali.Requery
Do Until RSKembali.EOF
    Combo6.AddItem RSKembali!TanggalKbl
    Combo7.AddItem Format(RSKembali!TanggalKbl, "YYYY ,MM, DD")
    Combo8.AddItem Format(RSKembali!TanggalKbl, "YYYY ,MM, DD")
    RSKembali.MoveNext
Loop

For i = 1 To 12
    Combo9.AddItem i
Next i
For i = 1 To 10
    Combo10.AddItem 2000 + i
Next i

End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
If Combo1 = "" Or Keyascii = 27 Then Unload Me
End Sub

'laporan peminjaman

'Lap Harian
Private Sub Combo1_Click()
    CR.SelectionFormula = "Totext({Pinjam.TanggalPjm})='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\Lap Pinjam Harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo2_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    If Combo2 = "" Then
        MsgBox "TanggalPjm awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    End If
    CR.SelectionFormula = "{Pinjam.TanggalPjm} in date (" & Combo2.Text & ") to date (" & Combo3.Text & ")"
    CR.ReportFileName = App.Path & "\Lap Pinjam Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo4_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Bulanan
Private Sub Combo5_Click()
    Call BukaDB
    RSPinjam.Open "select * from Pinjam where month(TanggalPjm)='" & Val(Combo4) & "' and year(TanggalPjm)='" & (Combo5) & "'", Conn
    If RSPinjam.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    
    CR.SelectionFormula = "Month({Pinjam.TanggalPjm})=" & Val(Combo4.Text) & " and Year({Pinjam.TanggalPjm})=" & Val(Combo5.Text)
    CR.ReportFileName = App.Path & "\Lap Pinjam Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

' laporan pengembalian

Private Sub Combo6_Keypress(Keyascii As Integer)
If Combo6 = "" Or Keyascii = 27 Then Unload Me
End Sub

'Lap Harian
Private Sub Combo6_Click()
    CR.SelectionFormula = "Totext({Kembali.TanggalKbl})='" & Combo6 & "'"
    CR.ReportFileName = App.Path & "\Lap Kembali Harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub combo7_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub combo8_Click()
    If Combo7 = "" Then
        MsgBox "TanggalKbl awal kosong", , "Informasi"
        Combo7.SetFocus
        Exit Sub
    End If
    CR.SelectionFormula = "{Kembali.TanggalKbl} in date (" & Combo7.Text & ") to date (" & Combo8.Text & ")"
    CR.ReportFileName = App.Path & "\Lap Kembali Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub combo9_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Bulanan
Private Sub combo10_Click()
    Call BukaDB
    RSKembali.Open "select * from Kembali where month(TanggalKbl)='" & Val(Combo9) & "' and year(TanggalKbl)='" & (Combo10) & "'", Conn
    If RSKembali.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo9.SetFocus
    End If
    
    CR.SelectionFormula = "Month({Kembali.TanggalKbl})=" & Val(Combo9.Text) & " and Year({Kembali.TanggalKbl})=" & Val(Combo10.Text)
    CR.ReportFileName = App.Path & "\Lap Kembali Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

