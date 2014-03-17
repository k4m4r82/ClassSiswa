Attribute VB_Name = "modDatabase"
'05 Agustus 2008

Option Explicit

Public Function konekToServer() As Boolean
    Dim strCon As String
    
    On Error GoTo errHandle
    
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\sampleDB.mdb"
    Set conn = New ADODB.Connection
    conn.ConnectionString = strCon
    conn.Open
    
    konekToServer = True
    
    Exit Function
errHandle:
    konekToServer = False
End Function

Public Sub closeRecordset(ByVal vRs As ADODB.Recordset)
    On Error Resume Next
    
    If Not (vRs Is Nothing) Then
        If vRs.State = adStateOpen Then vRs.Close
    End If
    
    Set vRs = Nothing
End Sub

Public Function openRecordset(ByVal query As String) As ADODB.Recordset
    Dim obj As ADODB.Recordset
    
    Set obj = New ADODB.Recordset
    obj.CursorLocation = adUseClient
    obj.Open query, conn, adOpenForwardOnly, adLockReadOnly
    
    Set openRecordset = obj
End Function

Public Function getRecordCount(ByVal vRs As ADODB.Recordset) As Long
''    On Error Resume Next

    vRs.MoveLast
    getRecordCount = vRs.RecordCount
    vRs.MoveFirst
End Function

