Option Compare Database
Option Explicit

Private dic As Object   ' 辞書

'---------------------------------------------
' 存在確認
'---------------------------------------------
Public Function Exists(ByVal kkey As Variant) As Boolean
    Exists = dic.Exists(kkey)
End Function
'---------------------------------------------
' プロパティにしてみた、、
'---------------------------------------------
Property Get count() As Long
    count = dic.count
End Property
'---------------------------------------------
'---------------------------------------------
Private Sub Class_Initialize()
    Set dic = CreateObject("Scripting.Dictionary")
End Sub
'---------------------------------------------
'---------------------------------------------
Private Sub Class_Terminate()
    Set dic = Nothing
End Sub
'---------------------------------------------
'---------------------------------------------
Public Sub init()
    dic.RemoveAll
End Sub
'---------------------------------------------
'---------------------------------------------
Public Sub add(ByVal kkey As Variant, ByVal dat As Variant)
    dic.add kkey, dat
End Sub
'---------------------------------------------
'---------------------------------------------
Public Function getDat(ByVal kkey As Variant) As Variant
    getDat = dic.Item(kkey)
End Function
'---------------------------------------------
'---------------------------------------------
Public Function setDat(ByVal kkey As Variant, ByVal dat As Variant)
    dic.Item(kkey) = dat
End Function
'---------------------------------------------
'---------------------------------------------
Public Function setObj(ByVal kkey As Variant, ByVal dat As Object)
    Set dic.Item(kkey) = dat
End Function
'---------------------------------------------
'---------------------------------------------
Public Function getObjNum(ByVal no As Long) As Object
    Dim a As Object
    a = dic.keys
    Set getObjNum = getObj(a(no))
End Function
'---------------------------------------------
'---------------------------------------------
Public Function getDatNum(ByVal no As Long) As Variant
    Dim a As Variant
    a = dic.keys
    getDatNum = getDat(a(no))
End Function
'---------------------------------------------
' key name を返す
'---------------------------------------------
Public Function getKey(ByVal no As Long) As Variant
    Dim a As Variant
    a = dic.keys
    getKey = a(no)
End Function
'---------------------------------------------
'---------------------------------------------
Public Sub addObj(ByVal kkey As Variant, ByVal dat As Object)
    dic.add kkey, dat
End Sub
'---------------------------------------------
'---------------------------------------------
Public Function getObj(kkey As Variant) As Object
    Set getObj = dic.Item(kkey)
End Function
