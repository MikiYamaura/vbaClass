Option Compare Database
Option Explicit

Private conn As ADODB.connection
'---------------------------------------------
'---------------------------------------------
Private Sub Class_Initialize()
    Set conn = CurrentProject.connection
End Sub
'---------------------------------------------
' DBÇêÿÇËë÷Ç¶ÇÁÇÍÇÈÇÊÇ§Ç…ÇµÇƒÇ®Ç≠
'---------------------------------------------
Property Let connection(con As ADODB.connection)
    Set conn = con
End Property

Property Get connection() As ADODB.connection
    Set connection = conn
End Property
'---------------------------------------------
'---------------------------------------------
Public Function getSqlValue(sql As String) As Variant
    Dim v As Variant
    Dim rs As ADODB.Recordset
    
    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        v = rs(0).value
    Else
        v = Null
    End If
    getSqlValue = v
End Function
'---------------------------------------------
'---------------------------------------------
Public Sub execSql(sql As String)
    On Error GoTo Erx
    
    conn.Execute (sql)
    
    Exit Sub
Erx:
    MsgBox ("SQL Error:" & Err.Description & Chr$(13) & Chr$(10) & "sql:" & sql)
    Stop
End Sub
'---------------------------------------------
'---------------------------------------------
Public Sub execSqlList(sql() As String)
    Dim i As Integer
    
    For i = LBound(sql) To UBound(sql)
        If sql(i) <> "" Then
            execSql (sql(i))
        End If
    Next

End Sub
'---------------------------------------------
'---------------------------------------------
'Private Sub Class_Terminate()
'End Sub

