VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LispObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' LispObj - every heap value of the Lisp interpreter: pairs, symbols,
' procedures, environments, vectors, hash tables, records, promises, ...
' It carries a Car/Cdr pair, a growable item array and an insertion-ordered
' hash table (keys are either strings or object identities).
' ---------------------------------------------------------------------------
Public Kind As Long         ' LK_* kind
Public Id As Long           ' unique id (identity hash)
Public SName As String      ' symbol name / procedure name / message
Public H As Long            ' hash of the symbol name
Public Code As Long         ' special-form code (symbols) / primitive code
Public Flags As Long
Public Ln As Long           ' source line (pairs produced by the reader)
Public Car As Variant
Public Cdr As Variant
Public Env As LispObj       ' closure environment / parent environment
Public Ref As LispObj       ' rest parameter / record type / misc
Public I1 As Long
Public I2 As Long

' ---- item array ----
Public Count As Long
Private mItems() As Variant
Private mCap As Long

' ---- hash table ----
Public DCount As Long
Private eN As Long
Private eKey() As Variant
Private eVal() As Variant
Private eKS() As String
Private eH() As Long
Private eLive() As Boolean
Private slots() As Long
Private sMask As Long

' field setters (an object property passed ByRef to another procedure is a
' temporary copy, so writes must happen inside the class)
Public Sub SetCar(ByRef v As Variant)
    If IsObject(v) Then Set Car = v Else Car = v
End Sub

Public Sub SetCdr(ByRef v As Variant)
    If IsObject(v) Then Set Cdr = v Else Cdr = v
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

Public Sub SetI(ByVal i As Long, ByRef v As Variant)
    If IsObject(v) Then
        Set mItems(i) = v
    Else
        mItems(i) = v
    End If
End Sub

Public Sub ClearItems()
    If mCap > 0 Then
        Erase mItems
    End If
    mCap = 0
    Count = 0
End Sub

' =========================== hash table ====================================
Private Sub InitSlots(ByVal nslots As Long)
    Dim i As Long
    ReDim slots(0 To nslots - 1)
    For i = 0 To nslots - 1
        slots(i) = -1
    Next
    sMask = nslots - 1
End Sub

' string-keyed lookup
Public Function FindH(ByRef ks As String, ByVal h As Long) As Long
    Dim p As Long, e As Long
    FindH = -1
    If sMask = 0 Then Exit Function
    p = h And sMask
    Do
        e = slots(p)
        If e < 0 Then Exit Function
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

' identity-keyed lookup (environments: keys are interned symbols)
Public Function FindRef(ByVal o As LispObj, ByVal h As Long) As Long
    Dim p As Long, e As Long
    FindRef = -1
    If sMask = 0 Then Exit Function
    p = h And sMask
    Do
        e = slots(p)
        If e < 0 Then Exit Function
        If eLive(e) Then
            If eH(e) = h Then
                If IsObject(eKey(e)) Then
                    If eKey(e) Is o Then
                        FindRef = e
                        Exit Function
                    End If
                End If
            End If
        End If
        p = (p + 1) And sMask
    Loop
End Function

Private Sub AddEntry(ByRef ks As String, ByVal h As Long, ByRef key As Variant, ByRef v As Variant)
    Dim p As Long, nc As Long
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

Public Sub SetH(ByRef ks As String, ByVal h As Long, ByRef key As Variant, ByRef v As Variant)
    Dim e As Long
    e = FindH(ks, h)
    If e >= 0 Then
        If IsObject(v) Then
            Set eVal(e) = v
        Else
            eVal(e) = v
        End If
    Else
        AddEntry ks, h, key, v
    End If
End Sub

Public Sub SetRef(ByVal o As LispObj, ByVal h As Long, ByRef v As Variant)
    Dim e As Long, k As Variant
    e = FindRef(o, h)
    If e >= 0 Then
        If IsObject(v) Then
            Set eVal(e) = v
        Else
            eVal(e) = v
        End If
    Else
        Set k = o
        AddEntry vbNullString, h, k, v
    End If
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

' break all references (frees reference cycles after a run)
Public Sub Release()
    Set Env = Nothing
    Set Ref = Nothing
    Car = Empty
    Cdr = Empty
    ClearItems
    DClear
End Sub
