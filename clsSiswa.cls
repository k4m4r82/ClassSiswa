VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRs             As ADODB.Recordset

Private mNomorInduk     As String 'm untuk prefix Member
Private mNama           As String
Private mAlamat         As String

'untuk mengeset nilai properties
'jika kode ini tidak kita buat maka sifat dari propertiesnya adalah readonly
Public Property Let nomorInduk(ByVal vData As String)
    mNomorInduk = vData
End Property

'untuk membaca nilai properties
Public Property Get nomorInduk() As String
    nomorInduk = mNomorInduk
End Property

Public Property Let nama(ByVal vData As String)
    mNama = vData
End Property
Public Property Get nama() As String
    nama = mNama
End Property

Public Property Let alamat(ByVal vData As String)
    mAlamat = vData
End Property
Public Property Get alamat() As String
    alamat = mAlamat
End Property

Public Function addData() As Boolean
    On Error GoTo errHandle
    
    strSql = "INSERT INTO siswa (nomorinduk, nama, alamat) VALUES ('" & _
             rep(nomorInduk) & "', '" & rep(nama) & "', '" & rep(alamat) & "')"
    conn.Execute strSql
    
    addData = True
    
    Exit Function
errHandle:
    addData = False
End Function

Public Function editData() As Boolean
    On Error GoTo errHandle
    
    strSql = "UPDATE siswa SET nama = '" & rep(nama) & "', alamat = '" & rep(alamat) & "' " & _
             "WHERE nomorinduk = '" & rep(nomorInduk) & "'"
    conn.Execute strSql
    
    editData = True
    
    Exit Function
errHandle:
    editData = False
End Function


Public Function delData() As Boolean
    On Error GoTo errHandle
    
    strSql = "DELETE FROM siswa WHERE nomorinduk = '" & rep(nomorInduk) & "'"
    conn.Execute strSql
    
    delData = True
    
    Exit Function
errHandle:
    delData = False
End Function

Public Function startGetData() As Long
'    On Error GoTo errHandle

    strSql = "SELECT nomorinduk, nama, alamat " & _
             "FROM siswa " & _
             "ORDER BY nama"
    Set mRs = openRecordset(strSql)
    If Not mRs.EOF Then startGetData = getRecordCount(mRs)
    
    Exit Function
errHandle:
    startGetData = 0
End Function

Public Sub getData()
'    On Error Resume Next
    
    mNomorInduk = IIf(IsNull(mRs("nomorinduk").Value), "", mRs("nomorinduk").Value)
    mNama = IIf(IsNull(mRs("nama").Value), "", mRs("nama").Value)
    mAlamat = IIf(IsNull(mRs("alamat").Value), "", mRs("alamat").Value)
    
    mRs.MoveNext
End Sub

'setelah pemanggilan method startGetData
'hasil recordset sebaiknya dihapus dari memory
Public Sub endGetData()
    Call closeRecordset(mRs)
End Sub

