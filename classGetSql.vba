Option Compare Database
Option Explicit

Private Table As New classDic
Private tmpRecord As classDic  ' どこかのレコードの内容を保持する
Private recNo As Long

Private idx As New classDic

'---------------------------------------------
' Index を作るけど、、 重複したら、スキップする
'---------------------------------------------
Public Sub makeIdx(colName As String)
    Dim rec As Long
    Dim dat As classDic
    Dim kkey As Variant
    
    idx.init
    For rec = 0 To Table.count - 1
        Set dat = Table.getObj(rec)
        kkey = dat.getDat(colName)
        If Not idx.Exists(kkey) Then
            idx.add kkey, rec
        End If
    Next
End Sub
'---------------------------------------------
' Index によるレコード取得
'---------------------------------------------
Public Function getIdx(kkey As Variant) As classDic
    Dim rec As Long
    
    If idx.Exists(kkey) Then
        recNo = idx.getDat(kkey)
        Set tmpRecord = getRecord(recNo)
    Else
        Set tmpRecord = Nothing
    End If
    Set getIdx = tmpRecord
End Function
'---------------------------------------------
'---------------------------------------------
Private Sub Class_Initialize()
    Set tmpRecord = Nothing
    recNo = 0
End Sub
'---------------------------------------------
'---------------------------------------------
Private Sub Class_Terminate()
    Table.init
    Set Table = Nothing
    Set tmpRecord = Nothing
End Sub
'---------------------------------------------
' レコードの内容取得
'---------------------------------------------
Property Get Record() As classDic
    Set Record = tmpRecord
End Property
'---------------------------------------------
' レコード番号
'---------------------------------------------
Property Get RecordNo() As Long
    RecordNo = recNo
End Property
'---------------------------------------------
' 初めのレコード
'---------------------------------------------
Public Sub MoveFirst()
    recNo = 0
    Set tmpRecord = getRecord(recNo)
End Sub
'---------------------------------------------
'---------------------------------------------
Public Sub MoveNext()
    recNo = recNo + 1
    Set tmpRecord = getRecord(recNo)
End Sub
'---------------------------------------------
'---------------------------------------------
Property Get EOF()
    Dim ret As Boolean
    If recNo >= Table.count Then
        ret = True
    Else
        ret = False
    End If
    EOF = ret
End Sub
'---------------------------------------------
' 内容
'---------------------------------------------
Public Function getDatNum(ByVal rec As Long, ByVal col As Long) As Variant
    Dim recDat As classDic
    Dim dat As Variant

    dat = Null
    Set recDat = getRecord(rec)

    If Not recDat Is Nothing Then
        If col < recDat.count Then
            dat = recDat.getDatNum(col)
        End If
    End If
    getDatNum = dat
End Function
'---------------------------------------------
' 内容(列名指定)
'---------------------------------------------
Public Function getDat(ByVal rec As Long, ByVal colName As String) As Variant
    Dim recDat As classDic
    Dim dat As Variant

    dat = Null
    Set recDat = getRecord(rec)

    If Not recDat Is Nothing Then
        dat = recDat.getDat(colName)
    End If
    getDat = dat
End Function
'---------------------------------------------
'---------------------------------------------
Public Function getRecord(ByVal rec As Long) As classDic
    Dim recDat As classDic

    Set recDat = Nothing
    If Table.count > rec Then
        Set recDat = Table.getObj(rec)
    End If
    Set getRecord = recDat
End Function
'---------------------------------------------
' 列名
'---------------------------------------------
Public Function getColName(ByVal col As Long) As String
    Dim nm As String

    nm = ""
    If Table.count > 0 Then
        If col < tmpRecord.count Then
            nm = tmpRecord.getKey(col)
        End If
    End If
    getColName = nm
End Function
'---------------------------------------------
' レコード数
'---------------------------------------------
Property Get recCount() As Long
    recCount = Table.count
End Property
'---------------------------------------------
' 列数
'---------------------------------------------
Property Get colCount() As Long
    Dim cnt As Long

    cnt = 0
    If Table.count > 0 Then
        cnt = tmpRecord.count
    End If
    colCount = cnt
End Property
'---------------------------------------------
' sql を実行してその全レコードの要素を取り込む
'---------------------------------------------
Public Sub getSql(ByVal sql As String)
    Dim cSql As New classSql
    Dim rs As New ADODB.Recordset
    Dim col As Long

    Set tmpRecord = Nothing
    Table.init
    rs.Open sql, cSql.connection, adOpenForwardOnly, adLockReadOnly
    recNo = 0
    col = 0
    Do While Not rs.EOF
        Set tmpRecord = New classDic
        For col = 0 To (rs.Fields.count - 1)
            tmpRecord.add rs.Fields(col).Name, rs(col).value
        Next
        Table.addObj recNo, tmpRecord
        recNo = recNo + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

