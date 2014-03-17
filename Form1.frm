VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lsvSiswa 
      Height          =   4695
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   2670
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdPerbaiki 
      Caption         =   "Perbaiki"
      Height          =   375
      Left            =   1455
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With lsvSiswa
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        
        .ColumnHeaders.Add , , "No.", 500
        .ColumnHeaders.Add , , "Nomor Induk", 1200
        .ColumnHeaders.Add , , "Nama", 3000
        .ColumnHeaders.Add , , "Nilai", 700
    End With
End Sub
