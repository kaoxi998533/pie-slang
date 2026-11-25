VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PyObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' PyObj - every heap value of the interpreter (list, tuple, dict, set,
' function, class, instance, module, scope, iterator, ...).
' A PyObj carries both a growable item array (lists, tuples, class bases)
' and an insertion-ordered hash table (dicts, sets, scopes, attributes).
' ---------------------------------------------------------------------------
Public Kind As Long         ' OK_* object kind
Public Id As Long           ' unique id (identity hash)
Public PName As String      ' name of function / class / module / method
Public Code As Long         ' builtin code / iterator kind
Public Flags As Long        ' misc flags
Public Node As PyNode       ' def / lambda node of a function
Public Env As PyObj         ' closure scope / parent scope
Public Cls As PyObj         ' class of an instance / defining class
Public Ref As PyObj         ' iterator source, property etc.
Public Defaults As PyObj    ' default argument values of a function
Public Decl As PyObj        ' scope: global / nonlocal declarations
Public SelfV As Variant     ' bound self / misc value
Public FnV As Variant       ' bound function / default factory / misc
Public I1 As Long           ' integer slots (range / iterator / traceback)
Public I2 As Long
Public I3 As Long

' ---- item array ----
Public Count As Long
Private mItems() As Variant
Private mCap As Long
Private mStr As String

' ---- hash table ----
Public DCount As Long       ' number of live entries
Private eN As Long          ' used entry slots (live + deleted)
Private eKey() As Variant
Private eVal() As Variant
Private eKS() As String
Private eH() As Long
Private eLive() As Boolean
Private slots() As Long
Private sMask As Long

' field setters (an object property passed ByRef to another procedure is a
' temporary copy, so writes must happen inside the class)
Public Sub SetFnV(ByRef v As Variant)
    If IsObject(v) Then Set FnV = v Else FnV = v
End Sub

Public Sub SetSelfV(ByRef v As Variant)
    If IsObject(v) Then Set SelfV = v Else SelfV = v
End Sub

' =========================== item array ====================================
Public Sub Reserve(ByVal n As Long)
    Dim nc As Long
    If n <= mCap Then Exit Sub
    nc = mCap * 2
    If nc < 4 Then nc = 4
    If nc < n Then nc = n
    If mCap = 0 Then
        ReDim mItems(0 To nc - 1)
    Else
        ReDim Preserve mItems(0 To nc - 1)
    End If
    mCap = nc
End Sub

Public Sub Push(ByRef v As Variant)
    If Count >= mCap Then Reserve Count + 1
    If IsObject(v) Then
        Set mItems(Count) = v
    Else
        mItems(Count) = v
    End If
    Count = Count + 1
End Sub

Public Sub GetInto(ByVal i As Long, ByRef out As Variant)
    If IsObject(mItems(i)) Then
        Set out = mItems(i)
    Else
        out = mItems(i)
    End If
End Sub

Public Function ItemIsObj(ByVal i As Long) As Boolean
    ItemIsObj = IsObject(mItems(i))
End Function

Public Function ItemType(ByVal i As Long) As Long
    ItemType = VarType(mItems(i))
End Function

Public Sub SetI(ByVal i As Long, ByRef v As Variant)
    If IsObject(v) Then
        Set mItems(i) = v
    Else
        mItems(i) = v
    End If
End Sub

Public Sub InsertAt(ByVal i As Long, ByRef v As Variant)
    Dim j As Long
    If Count >= mCap Then Reserve Count + 1
    For j = Count To i + 1 Step -1
        If IsObject(mItems(j - 1)) Then
            Set mItems(j) = mItems(j - 1)
        Else
            mItems(j) = mItems(j - 1)
        End If
    Next
    If IsObject(v) Then
        Set mItems(i) = v
    Else
        mItems(i) = v
    End If
    Count = Count + 1
End Sub

Public Sub RemoveAt(ByVal i As Long)
    Dim j As Long
    For j = i To Count - 2
        If IsObject(mItems(j + 1)) Then
            Set mItems(j) = mItems(j + 1)
        Else
            mItems(j) = mItems(j + 1)
        End If
    Next
    Count = Count - 1
    mItems(Count) = Empty
End Sub

Public Sub ClearItems()
    If mCap > 0 Then
        Erase mItems
    End If
    mCap = 0
    Count = 0
End Sub

Public Sub GetArray(ByRef arr() As Variant)
    Dim i As Long
    If Count = 0 Then
        ReDim arr(0 To 0)
        Exit Sub
    End If
    ReDim arr(0 To Count - 1)
    For i = 0 To Count - 1
        If IsObject(mItems(i)) Then
            Set arr(i) = mItems(i)
        Else
            arr(i) = mItems(i)
        End If
    Next
End Sub

Public Sub SetArray(ByRef arr() As Variant, ByVal n As Long)
    Dim i As Long
    ClearItems
    Reserve n
    For i = 0 To n - 1
        If IsObject(arr(i)) Then
            Set mItems(i) = arr(i)
        Else
            mItems(i) = arr(i)
        End If
    Next
    Count = n
End Sub

Public Sub ReverseItems()
    Dim i As Long, j As Long, t As Variant
    i = 0
    j = Count - 1
    Do While i < j
        If IsObject(mItems(i)) Then
            Set t = mItems(i)
        Else
            t = mItems(i)
        End If
        If IsObject(mItems(j)) Then
            Set mItems(i) = mItems(j)
        Else
            mItems(i) = mItems(j)
        End If
        If IsObject(t) Then
            Set mItems(j) = t
        Else
            mItems(j) = t
        End If
        i = i + 1
        j = j - 1
    Loop
End Sub

' string payload (used by string iterators)
Public Sub SetStr(ByVal s As String)
    mStr = s
End Sub

' one step of a built-in iterator; returns False when exhausted
Public Function IterStep(ByRef out As Variant) As Boolean
    Select Case Code
    Case 1  ' sequence (list / tuple), live index
        If I1 < Ref.Count Then
            Ref.GetInto I1, out
            I1 = I1 + 1
            IterStep = True
        End If
    Case 2  ' string
        If I1 < I2 Then
            I1 = I1 + 1
            out = Mid$(mStr, I1, 1)
            IterStep = True
        End If
    Case 3  ' range
        If I3 > 0 Then
            If I1 >= I2 Then Exit Function
        Else
            If I1 <= I2 Then Exit Function
        End If
        out = I1
        IterStep = True
        If CDbl(I1) + CDbl(I3) > 2147483647# Or CDbl(I1) + CDbl(I3) < -2147483648# Then
            I1 = I2
        Else
            I1 = I1 + I3
        End If
    End Select
End Function

' ============================ hash table ===================================
Private Sub InitSlots(ByVal nslots As Long)
    Dim i As Long
    ReDim slots(0 To nslots - 1)
    For i = 0 To nslots - 1
        slots(i) = -1
    Next
    sMask = nslots - 1
End Sub

Public Function FindH(ByRef ks As String, ByVal h As Long) As Long
    Dim p As Long, e As Long
    If sMask = 0 Then
        FindH = -1
        Exit Function
    End If
    p = h And sMask
    Do
        e = slots(p)
        If e < 0 Then
            FindH = -1
            Exit Function
        End If
        If eLive(e) Then
            If eH(e) = h Then
                If eKS(e) = ks Then
                    FindH = e
                    Exit Function
                End If
            End If
        End If
        p = (p + 1) And sMask
    Loop
End Function

Public Sub SetH(ByRef ks As String, ByVal h As Long, ByRef key As Variant, ByRef v As Variant)
    Dim e As Long, p As Long, nc As Long
    e = FindH(ks, h)
    If e >= 0 Then
        If IsObject(v) Then
            Set eVal(e) = v
        Else
            eVal(e) = v
        End If
        Exit Sub
    End If
    If sMask = 0 Then
        InitSlots 8
        ReDim eKey(0 To 7)
        ReDim eVal(0 To 7)
        ReDim eKS(0 To 7)
        ReDim eH(0 To 7)
        ReDim eLive(0 To 7)
    ElseIf (eN + 1) * 2 > sMask + 1 Then
        Rehash
    End If
    If eN > UBound(eKS) Then
        nc = (UBound(eKS) + 1) * 2
        ReDim Preserve eKey(0 To nc - 1)
        ReDim Preserve eVal(0 To nc - 1)
        ReDim Preserve eKS(0 To nc - 1)
        ReDim Preserve eH(0 To nc - 1)
        ReDim Preserve eLive(0 To nc - 1)
    End If
    eKS(eN) = ks
    eH(eN) = h
    If IsObject(key) Then
        Set eKey(eN) = key
    Else
        eKey(eN) = key
    End If
    If IsObject(v) Then
        Set eVal(eN) = v
    Else
        eVal(eN) = v
    End If
    eLive(eN) = True
    p = h And sMask
    Do While slots(p) >= 0
        p = (p + 1) And sMask
    Loop
    slots(p) = eN
    eN = eN + 1
    DCount = DCount + 1
End Sub

Private Sub Rehash()
    Dim i As Long, j As Long, ns As Long, p As Long
    j = 0
    For i = 0 To eN - 1
        If eLive(i) Then
            If i <> j Then
                eKS(j) = eKS(i)
                eH(j) = eH(i)
                If IsObject(eKey(i)) Then
                    Set eKey(j) = eKey(i)
                Else
                    eKey(j) = eKey(i)
                End If
                If IsObject(eVal(i)) Then
                    Set eVal(j) = eVal(i)
                Else
                    eVal(j) = eVal(i)
                End If
                eLive(j) = True
            End If
            j = j + 1
        End If
    Next
    For i = j To eN - 1
        eLive(i) = False
        eKey(i) = Empty
        eVal(i) = Empty
        eKS(i) = vbNullString
    Next
    eN = j
    ns = 8
    Do While ns < (eN + 1) * 4
        ns = ns * 2
    Loop
    InitSlots ns
    For i = 0 To eN - 1
        p = eH(i) And sMask
        Do While slots(p) >= 0
            p = (p + 1) And sMask
        Loop
        slots(p) = i
    Next
End Sub

Public Function DelH(ByRef ks As String, ByVal h As Long) As Boolean
    Dim e As Long
    e = FindH(ks, h)
    If e < 0 Then Exit Function
    eLive(e) = False
    eKey(e) = Empty
    eVal(e) = Empty
    DCount = DCount - 1
    DelH = True
End Function

Public Sub DelAt(ByVal e As Long)
    If eLive(e) Then
        eLive(e) = False
        eKey(e) = Empty
        eVal(e) = Empty
        DCount = DCount - 1
    End If
End Sub

Public Function EntryCount() As Long
    EntryCount = eN
End Function

Public Function LiveAt(ByVal i As Long) As Boolean
    LiveAt = eLive(i)
End Function

Public Sub KeyInto(ByVal i As Long, ByRef out As Variant)
    If IsObject(eKey(i)) Then
        Set out = eKey(i)
    Else
        out = eKey(i)
    End If
End Sub

Public Sub ValInto(ByVal i As Long, ByRef out As Variant)
    If IsObject(eVal(i)) Then
        Set out = eVal(i)
    Else
        out = eVal(i)
    End If
End Sub

Public Function ValType(ByVal i As Long) As Long
    ValType = VarType(eVal(i))
End Function

Public Sub SetValAt(ByVal i As Long, ByRef v As Variant)
    If IsObject(v) Then
        Set eVal(i) = v
    Else
        eVal(i) = v
    End If
End Sub

Public Function KSAt(ByVal i As Long) As String
    KSAt = eKS(i)
End Function

Public Function HAt(ByVal i As Long) As Long
    HAt = eH(i)
End Function

Public Sub DClear()
    If sMask <> 0 Then
        Erase eKey
        Erase eVal
        Erase eKS
        Erase eH
        Erase eLive
        Erase slots
    End If
    sMask = 0
    eN = 0
    DCount = 0
End Sub

' break all references (used to free reference cycles after a run)
Public Sub Release()
    Set Node = Nothing
    Set Env = Nothing
    Set Cls = Nothing
    Set Ref = Nothing
    Set Defaults = Nothing
    Set Decl = Nothing
    SelfV = Empty
    FnV = Empty
    mStr = vbNullString
    ClearItems
    DClear
End Sub
