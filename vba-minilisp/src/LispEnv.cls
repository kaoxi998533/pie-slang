VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LispEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' LispEnv —— 环境（变量表）
'
' 一个环境 = 一张“名字 -> 值”的表 + 一个指向外层环境的 Parent。
'
'   全局环境:  + -> #<builtin +>,  x -> 10, ...      (Parent = Nothing)
'       ^
'       | Parent
'   调用 (f 1 2) 时新建的环境:  a -> 1,  b -> 2
'
' 查找变量时先找自己这一层，找不到就去 Parent 里找，一直到全局环境。
' 这就是“作用域”和“闭包”的全部秘密。
'
' 为了简单，这里用两个数组线性查找；想更快可以换成 Collection 或
' Scripting.Dictionary（见 README 练习）。
'==============================================================
Option Explicit

Public Parent As LispEnv        ' 外层环境；全局环境的 Parent 是 Nothing

Private mNames() As String      ' 变量名
Private mValues() As Variant    ' 对应的值
Private mCount As Long          ' 已经存了几个变量

Private Sub Class_Initialize()
    ReDim mNames(0 To 7)
    ReDim mValues(0 To 7)
    mCount = 0
    Set Parent = Nothing          ' VBA 里本来就是 Nothing；LibreOffice 需要明确写出来
End Sub

' 在“本层”找名字，返回下标；找不到返回 -1
Private Function IndexOf(ByVal nm As String) As Long
    Dim i As Long
    For i = 0 To mCount - 1
        If mNames(i) = nm Then
            IndexOf = i
            Exit Function
        End If
    Next
    IndexOf = -1
End Function

' 本层有没有这个名字
Public Function Has(ByVal nm As String) As Boolean
    Has = (IndexOf(nm) >= 0)
End Function

' 取本层某个名字的值（调用前先用 Has 或 Find 确认它存在）
Public Function GetValue(ByVal nm As String) As Variant
    GetValue = mValues(IndexOf(nm))
End Function

' 在本层定义变量；已经有了就覆盖
Public Sub Define(ByVal nm As String, ByVal v As Variant)
    Dim i As Long
    i = IndexOf(nm)
    If i < 0 Then
        If mCount > UBound(mNames) Then
            ' 数组满了：容量翻倍
            ReDim Preserve mNames(0 To mCount * 2)
            ReDim Preserve mValues(0 To mCount * 2)
        End If
        i = mCount
        mNames(i) = nm
        mCount = mCount + 1
    End If
    mValues(i) = v
End Sub

' 从本层开始，沿着 Parent 一层层往外找，
' 返回“含有这个名字的那一层环境”；哪层都没有就返回 Nothing
Public Function Find(ByVal nm As String) As LispEnv
    Dim e As LispEnv
    Set e = Me
    Do While Not (e Is Nothing)
        If e.Has(nm) Then
            Set Find = e
            Exit Function
        End If
        Set e = e.Parent
    Loop
    Set Find = Nothing
End Function

' 清空本层（重新开始时调用，帮助 VBA 回收内存）
Public Sub Clear()
    ReDim mNames(0 To 7)
    ReDim mValues(0 To 7)
    mCount = 0
    Set Parent = Nothing
End Sub
