Attribute VB_Name = "modFunction"
Option Explicit

Public Function rep(ByVal kata As String) As String
    rep = Replace(kata, "'", "''")
End Function
