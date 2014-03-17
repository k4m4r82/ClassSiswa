VERSION 5.00
Begin VB.Form frmAddEditSiswa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Siswa"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1335
      Width           =   975
   End
   Begin VB.CommandButton cmdSelesai 
      Caption         =   "Selesai"
      Height          =   375
      Left            =   3255
      TabIndex        =   3
      Top             =   1335
      Width           =   975
   End
   Begin VB.TextBox txtAlamat 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   930
      Width           =   3135
   End
   Begin VB.TextBox txtNama 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   525
      Width           =   3135
   End
   Begin VB.TextBox txtNomorInduk 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   930
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Induk"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddEditSiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dataBaru As Boolean

Private Sub cmdSelesai_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    Dim objSiswa    As clsSiswa
    Dim row         As Long
    
    If dataBaru Then
        Set objSiswa = New clsSiswa
        With objSiswa
            .nomorInduk = txtNomorInduk.Text
            .nama = txtNama.Text
            .alamat = txtAlamat.Text
                    
            result = .addData
        End With
        Set objSiswa = Nothing
        
        If result Then
            With frmSiswa.lsvSiswa
                row = .ListItems.Count + 1
                
                .ListItems.Add , , row
                .ListItems(row).SubItems(1) = txtNomorInduk.Text
                .ListItems(row).SubItems(2) = txtNama.Text
                .ListItems(row).SubItems(3) = txtAlamat.Text
            End With
            
            txtNomorInduk.Text = ""
            txtNama.Text = ""
            txtAlamat.Text = ""
            txtNomorInduk.SetFocus
            
        Else
            MsgBox "Data siswa gagal disimpan", vbExclamation, "Peringatan"
        End If
        
    Else
        Set objSiswa = New clsSiswa
        With objSiswa
            .nomorInduk = txtNomorInduk.Text
            .nama = txtNama.Text
            .alamat = txtAlamat.Text
                    
            result = .editData
        End With
        Set objSiswa = Nothing
        
        If result Then
            With frmSiswa.lsvSiswa
                row = .SelectedItem.Index
                
                .ListItems(row).SubItems(1) = txtNomorInduk.Text
                .ListItems(row).SubItems(2) = txtNama.Text
                .ListItems(row).SubItems(3) = txtAlamat.Text
            End With
            
            Unload Me
            
        Else
            MsgBox "Data siswa gagal disimpan", vbExclamation, "Peringatan"
        End If
    End If
End Sub

