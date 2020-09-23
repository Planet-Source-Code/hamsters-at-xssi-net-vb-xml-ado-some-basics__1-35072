Attribute VB_Name = "ModADO"
Option Explicit

'*********************************************************
'* Make sure that you have the ADO 2.5 referenced
'*********************************************************

'Microsoft ActiveX Data Object Library 2.x

Public Conn1 As New ADODB.Connection

' above is the Connection used to coonect to the Database

' below is the Record Set

Public Rs1 As New ADODB.Recordset

'Database Connection Strings

Public StringCnn As String
Public AccessCnn As String

Public Function UpdateDB()
   ' this is the update db function. This is used when you add or edit a record
   ' Rs1!info = frmDE1.Text1.Text
    ' the rs1!info refer's to the info field in the selected recordset
    
    Rs1.UpdateBatch adAffectCurrent
    
End Function

Public Function AddDB()
    ' this is called to add a new row to the redorcset so you can add a new row.
    
    Rs1.AddNew
End Function

Public Function DeleteDB(DeleteField As String, DeleteValue As String)
    Rs1.Delete adAffectCurrent
    Rs1.UpdateBatch
End Function

Public Function OpenDB()


AccessCnn = "DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=axml.mdb;" & "DefaultDir=" & App.Path + "\" & ";" & "UID=admin;PWD=;"

 StringCnn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\axml.mdb;Persist Security Info=False;"


Conn1.ConnectionString = AccessCnn
Conn1.Open

Rs1.Open "SELECT * FROM [xml1]", StringCnn, adOpenStatic, adLockBatchOptimistic, adCmdText
End Function

Public Function CloseDB()
    On Error Resume Next

    Rs1.Close
    Conn1.Close
End Function



