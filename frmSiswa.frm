VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSiswa 
   Caption         =   "Data Siswa"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsvSiswa 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   2310
      TabIndex        =   2
      Top             =   4935
      Width           =   975
   End
   Begin VB.CommandButton cmdPerbaiki 
      Caption         =   "Perbaiki"
      Height          =   375
      Left            =   1215
      TabIndex        =   1
      Top             =   4935
      Width           =   975
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4935
      Width           =   975
   End
End
Attribute VB_Name = "frmSiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objSiswa    As clsSiswa
Dim row         As Long

Private Sub showDataSiswa()
    Dim i       As Long
    
    Set objSiswa = New clsSiswa
    With objSiswa
        Screen.MousePointer = vbHourglass
        DoEvents
        
        'saya disini menggunakan perulangan for untuk menampilkan data
        'biasanya kita untuk menampilkan data dari recordset menggunakan
        'perulangan : do while not varRs.EOF atau semisalnya
        'hasil survey membuktikan menggunakan for lebih cepat
        For i = 1 To .startGetData
            Call .getData
            
            lsvSiswa.ListItems.Add , , i
            lsvSiswa.ListItems(i).SubItems(1) = .nomorInduk
            lsvSiswa.ListItems(i).SubItems(2) = .nama
            lsvSiswa.ListItems(i).SubItems(3) = .alamat
        Next i
        Call .endGetData

        Screen.MousePointer = vbDefault
    End With
    Set objSiswa = Nothing
End Sub

Private Sub cmdHapus_Click()
    If MsgBox("Anda yakin ingin menghapus data ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
        row = lsvSiswa.SelectedItem.Index
                
        Set objSiswa = New clsSiswa
        With objSiswa
            .nomorInduk = lsvSiswa.ListItems(row).SubItems(1)
            result = .delData
            
            If result Then
                lsvSiswa.ListItems.Remove row
            Else
                MsgBox "Gagal menghapus data", vbExclamation, "Peringatan"
            End If
        End With
        Set objSiswa = Nothing
    End If
End Sub

Private Sub cmdPerbaiki_Click()
    row = lsvSiswa.SelectedItem.Index
    
    With frmAddEditSiswa
        .dataBaru = False
        
        .txtNomorInduk.Text = lsvSiswa.ListItems(row).SubItems(1)
        .txtNomorInduk.Enabled = False
        
        .txtNama.Text = lsvSiswa.ListItems(row).SubItems(2)
        .txtAlamat.Text = lsvSiswa.ListItems(row).SubItems(3)
        
        .Show vbModal
    End With
End Sub

Private Sub cmdTambah_Click()
    With frmAddEditSiswa
        .dataBaru = True
        .Show vbModal
    End With
End Sub

Private Sub Form_Load()
    With lsvSiswa
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        
        .ColumnHeaders.Add , , "No.", 500
        .ColumnHeaders.Add , , "Nomor Induk", 1200, lvwColumnCenter
        .ColumnHeaders.Add , , "Nama", 2000
        .ColumnHeaders.Add , , "Alamat", 4000
    End With
    
    Call showDataSiswa
End Sub
