VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form RincianKbl 
   Caption         =   "Rincian Pengembalian Buku"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3000
      Top             =   2040
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
      Height          =   1695
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
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
      ColumnCount     =   3
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "Denda"
         Caption         =   "Denda"
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
            ColumnWidth     =   2894.74
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1244.976
         EndProperty
      EndProperty
   End
   Begin VB.Label LblDenda 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   2280
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Denda"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   1920
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label LblTanggal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Anggota"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   1250
   End
   Begin VB.Label LblAnggota 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1250
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1920
      Width           =   750
   End
   Begin VB.Label LblJumlah 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   750
   End
End
Attribute VB_Name = "RincianKbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
Call BukaDB
List1.Clear
RSKembali.Open "Select Distinct NomorKbl from Kembali ", Conn
Do Until RSKembali.EOF
    List1.AddItem RSKembali!NomorKbl
    RSKembali.MoveNext
Loop
Conn.Close
End Sub

Private Sub list1_click()
Call BukaDB
Conn.CursorLocation = adUseClient
RSKembali.Open "select * from Kembali where NomorKbl='" & List1.Text & "'", Conn
RSKembali.Requery

If Not RSKembali.EOF Then LblTanggal = RSKembali!TanggalKbl

RSAnggota.Open "select * from Anggota where NomorAgt='" & RSKembali!NomorAgt & "'", Conn
If Not RSAnggota.EOF Then LblAnggota = RSAnggota!Namaagt
Conn.Close

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOVCD.mdb"
Adodc1.RecordSource = "select Judul, JumlahFlm as Jumlah,Denda from Film,detailKbl,Kembali where DetailKbl.NomorFlm=Film.NomorFlm and left(detailKbl.NomorKbl,8)=Kembali.NomorKbl and Kembali.NomorKbl='" & List1 & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
LblJumlah = Adodc1.Recordset.RecordCount
Call JmlDenda
End Sub

Private Sub List1_keyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Sub JmlDenda()
Adodc1.Recordset.MoveFirst
Denda = 0
Do While Not Adodc1.Recordset.EOF
    Denda = Denda + Adodc1.Recordset!Denda
    Adodc1.Recordset.MoveNext
Loop
LblDenda = Denda
End Sub
