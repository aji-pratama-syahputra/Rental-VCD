VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   4035
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "a"
            Object.ToolTipText     =   "Anggota"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "b"
            Object.ToolTipText     =   "Film"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "c"
            Object.ToolTipText     =   "Peminjaman"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "d"
            Object.ToolTipText     =   "Pengembalian"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "e"
            Object.ToolTipText     =   "Laporan Data Anggota"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "f"
            Object.ToolTipText     =   "Laporan Data Film"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "g"
            Object.ToolTipText     =   "Laporan Transaksi"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "h"
            Object.ToolTipText     =   "Rincian Peminjaman"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "i"
            Object.ToolTipText     =   "Rincian Pengembalian"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k"
            Object.ToolTipText     =   "Keluar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3540
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CR 
      Left            =   2880
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":0C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1BEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnanggota 
         Caption         =   "Anggota"
      End
      Begin VB.Menu mnfilm 
         Caption         =   "Film"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnpinjaman 
         Caption         =   "Pinjaman"
      End
      Begin VB.Menu mnkembali 
         Caption         =   "Pengembalian"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlapfilm 
         Caption         =   "Data Film"
      End
      Begin VB.Menu mnlapanggota 
         Caption         =   "Data Anggota"
      End
      Begin VB.Menu mnlaptransaksi 
         Caption         =   "Laporan Transaksi"
      End
   End
   Begin VB.Menu mnrincian 
      Caption         =   "Rincian"
      Begin VB.Menu mnrincianpjm 
         Caption         =   "Pinjaman"
      End
      Begin VB.Menu mnrinciankbl 
         Caption         =   "Pengembalian"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
pesan = MsgBox("Tutup aplikasi..?", vbYesNo)
If pesan = vbYes Then End
End If
End Sub

Private Sub mnanggota_Click()
Anggota.Show
End Sub

Private Sub mnFilm_Click()
Film.Show
End Sub

Private Sub mnkeluar_Click()
pesan = MsgBox("Tutup aplikasi..?", vbYesNo)
If pesan = vbYes Then End
End Sub

Private Sub mnkembali_Click()
Kembali.Show
End Sub

Private Sub mnlapanggota_Click()
CR.ReportFileName = App.Path & "\Lap Anggota.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 0
End Sub

Private Sub mnlapFilm_Click()
CR.ReportFileName = App.Path & "\Lap Film.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 0
End Sub

Private Sub mnlaptransaksi_Click()
Laporan.Show
End Sub

Private Sub mnpinjaman_Click()
Pinjam.Show
End Sub

Private Sub mnrinciankbl_Click()
RincianKbl.Show
End Sub

Private Sub mnrincianpjm_Click()
RincianPjm.Show
End Sub

Private Sub mnuji_Click()
UjiSQL.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "a"
        Anggota.Show
    Case "b"
        Film.Show
    Case "c"
        Pinjam.Show
    Case "d"
        Kembali.Show
    Case "e"
       CR.ReportFileName = App.Path & "\Lap anggota.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "f"
       CR.ReportFileName = App.Path & "\Lap Film.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "g"
        Laporan.Show
    Case "h"
        RincianPjm.Show
    Case "i"
        RincianKbl.Show
    Case "j"
        End
End Select

End Sub
