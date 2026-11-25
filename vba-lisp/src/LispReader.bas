Attribute VB_Name = "LispReader"
Option Explicit
' ===========================================================================
'  LispReader - turns source text into Lisp data (S-expressions).
'  Supports ( ) [ ], dotted pairs, ' ` , ,@, strings with escapes, #t/#f,
'  #\characters, #( vectors ), #x/#b/#o numbers, ; comments, #| block |#
'  comments and #; datum comments.  Pairs remember their source line.
' ===========================================================================

Private rS As String
Private rP As Long
Private rN As Long
Private rLine As Long

Private Sub ReadErr(ByVal msg As String)
    lgLine = rLine
    LThrow "read: " & msg
End Sub

' all data in the text, as a list
Public Function LReadAll(ByVal src As String) As Variant
    Dim head As LispObj, tail As LispObj, cell As LispObj, v As Variant, e As Variant
    rS = Replace(Replace(src, vbCrLf, vbLf), vbCr, vbLf)
    If Len(rS) > 0 Then
        If LUCode(Left$(rS, 1)) = 65279 Then rS = Mid$(rS, 2)
    End If
    rP = 1
    rN = Len(rS)
    rLine = 1
    Do
        SkipWs
        If rP > rN Then Exit Do
        ReadDatum v
        Set cell = LCons(v, e)
        cell.Ln = rLine
        If head Is Nothing Then
            Set head = cell
        Else
            Set tail.Cdr = cell
        End If
        Set tail = cell
    Loop
    If head Is Nothing Then
        LReadAll = Empty
    Else
        Set LReadAll = head
    End If
End Function

Private Function IsDelim(ByVal c As String) As Boolean
    Select Case c
    Case " ", vbTab, vbLf, vbCr, "(", ")", "[", "]", """", ";", "'", "`", ",", ChrW(160), vbFormFeed
        IsDelim = True
    End Select
End Function

Private Sub SkipWs()
    Dim c As String, depth As Long, v As Variant
    Do While rP <= rN
        c = Mid$(rS, rP, 1)
        If c = vbLf Then
            rLine = rLine + 1
            rP = rP + 1
        ElseIf c = " " Or c = vbTab Or c = vbCr Or c = vbFormFeed Or c = ChrW(160) Then
            rP = rP + 1
        ElseIf c = ";" Then
            Do While rP <= rN
                If Mid$(rS, rP, 1) = vbLf Then Exit Do
                rP = rP + 1
            Loop
        ElseIf c = "#" And Mid$(rS, rP + 1, 1) = "|" Then
            rP = rP + 2
            depth = 1
            Do While rP <= rN And depth > 0
                If Mid$(rS, rP, 2) = "|#" Then
                    depth = depth - 1
                    rP = rP + 2
                ElseIf Mid$(rS, rP, 2) = "#|" Then
                    depth = depth + 1
                    rP = rP + 2
                Else
                    If Mid$(rS, rP, 1) = vbLf Then rLine = rLine + 1
                    rP = rP + 1
                End If
            Loop
            If depth > 0 Then ReadErr "unterminated #| comment"
        ElseIf c = "#" And Mid$(rS, rP + 1, 1) = ";" Then
            rP = rP + 2
            ReadDatum v
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function ReadToken() As String
    Dim st As Long
    st = rP
    Do While rP <= rN
        If IsDelim(Mid$(rS, rP, 1)) Then Exit Do
        rP = rP + 1
    Loop
    ReadToken = Mid$(rS, st, rP - st)
End Function

Private Sub ReadDatum(ByRef out As Variant)
    Dim c As String, c2 As String, v As Variant, tok As String, ok As Boolean, ln As Long, e As Variant
    Dim lst As Variant, vec As LispObj, x As Variant, o As LispObj
    SkipWs
    If rP > rN Then ReadErr "unexpected end of input"
    c = Mid$(rS, rP, 1)
    ln = rLine
    Select Case c
    Case "(", "["
        rP = rP + 1
        If c = "(" Then
            ReadListTail ")", ln, out
        Else
            ReadListTail "]", ln, out
        End If
    Case ")", "]"
        ReadErr "unexpected '" & c & "'"
    Case "'"
        rP = rP + 1
        ReadDatum v
        Set o = LList2(lsQuote, v)
        o.Ln = ln
        Set out = o
    Case "`"
        rP = rP + 1
        ReadDatum v
        Set o = LList2(lsQuasi, v)
        o.Ln = ln
        Set out = o
    Case ","
        rP = rP + 1
        If Mid$(rS, rP, 1) = "@" Then
            rP = rP + 1
            ReadDatum v
            Set o = LList2(lsUnquoteSplicing, v)
        Else
            ReadDatum v
            Set o = LList2(lsUnquote, v)
        End If
        o.Ln = ln
        Set out = o
    Case """"
        out = ReadString()
    Case "#"
        c2 = Mid$(rS, rP + 1, 1)
        Select Case c2
        Case "("
            rP = rP + 2
            ReadListTail ")", ln, lst
            Set vec = LNewObj(LK_VECTOR)
            LSetV x, lst
            Do While LIsPair(x)
                Set o = x
                vec.Push o.Car
                LSetV x, o.Cdr
            Loop
            Set out = vec
        Case "\"
            rP = rP + 2
            Set out = ReadChar()
        Case "x", "X", "b", "B", "o", "O", "d", "D", "e", "E", "i", "I"
            rP = rP + 2
            tok = ReadToken()
            Select Case LCase$(c2)
            Case "x": v = LParseNumber(tok, 16, ok)
            Case "b": v = LParseNumber(tok, 2, ok)
            Case "o": v = LParseNumber(tok, 8, ok)
            Case Else: v = LParseNumber(tok, 10, ok)
            End Select
            If Not ok Then ReadErr "bad number #" & c2 & tok
            If LCase$(c2) = "i" And LIsExact(v) Then v = CDbl(v)
            out = v
        Case Else
            tok = ReadToken()
            Select Case tok
            Case "#t", "#true": out = True
            Case "#f", "#false": out = False
            Case "#!eof": Set out = lgEof
            Case "#!default", "#!unspecific", "#!void": Set out = lgVoid
            Case Else: ReadErr "unknown syntax " & tok
            End Select
        End Select
    Case Else
        tok = ReadToken()
        If Len(tok) = 0 Then ReadErr "unexpected character '" & c & "'"
        v = LParseNumber(tok, 10, ok)
        If ok Then
            out = v
        Else
            Set out = LIntern(tok)
        End If
    End Select
End Sub

Private Sub ReadListTail(ByVal closer As String, ByVal ln As Long, ByRef out As Variant)
    Dim head As LispObj, tail As LispObj, cell As LispObj, v As Variant, e As Variant, c As String
    Do
        SkipWs
        If rP > rN Then
            lgLine = ln
            LThrow "read: missing '" & closer & "' for the list starting at line " & ln
        End If
        c = Mid$(rS, rP, 1)
        If c = ")" Or c = "]" Then
            If c <> closer Then ReadErr "'" & c & "' does not match '" & IIf(closer = ")", "(", "[") & "'"
            rP = rP + 1
            Exit Do
        End If
        If c = "." And IsDelim(Mid$(rS, rP + 1, 1)) Then
            If head Is Nothing Then ReadErr "unexpected '.'"
            rP = rP + 1
            ReadDatum v
            If IsObject(v) Then
                Set tail.Cdr = v
            Else
                tail.Cdr = v
            End If
            SkipWs
            If Mid$(rS, rP, 1) <> closer Then ReadErr "expected '" & closer & "' after the dotted tail"
            rP = rP + 1
            Exit Do
        End If
        ReadDatum v
        Set cell = LCons(v, e)
        If head Is Nothing Then
            Set head = cell
            cell.Ln = ln
        Else
            Set tail.Cdr = cell
            cell.Ln = rLine
        End If
        Set tail = cell
    Loop
    If head Is Nothing Then
        out = Empty
    Else
        Set out = head
    End If
End Sub

Private Function ReadString() As String
    Dim parts() As String, np As Long, st As Long, c As String, hx As String, code As Long, ok As Boolean
    ReDim parts(0 To 15)
    rP = rP + 1
    st = rP
    Do
        If rP > rN Then ReadErr "unterminated string"
        c = Mid$(rS, rP, 1)
        If c = """" Then
            AddPart parts, np, Mid$(rS, st, rP - st)
            rP = rP + 1
            Exit Do
        ElseIf c = "\" Then
            AddPart parts, np, Mid$(rS, st, rP - st)
            c = Mid$(rS, rP + 1, 1)
            rP = rP + 2
            Select Case c
            Case "n": AddPart parts, np, vbLf
            Case "t": AddPart parts, np, vbTab
            Case "r": AddPart parts, np, vbCr
            Case "a": AddPart parts, np, ChrW(7)
            Case "0": AddPart parts, np, ChrW(0)
            Case "\": AddPart parts, np, "\"
            Case """": AddPart parts, np, """"
            Case "x", "u"
                hx = vbNullString
                Do While rP <= rN
                    c = Mid$(rS, rP, 1)
                    rP = rP + 1
                    If c = ";" Then Exit Do
                    hx = hx & c
                    If c = """" Or Len(hx) > 6 Then ReadErr "bad \x escape"
                Loop
                code = CLng(LParseInt(hx, 16, ok))
                If Not ok Then ReadErr "bad \x escape"
                AddPart parts, np, ChrW(code)
            Case vbLf
                rLine = rLine + 1
                Do While rP <= rN
                    c = Mid$(rS, rP, 1)
                    If c <> " " And c <> vbTab Then Exit Do
                    rP = rP + 1
                Loop
            Case Else
                AddPart parts, np, c
            End Select
            st = rP
        Else
            If c = vbLf Then rLine = rLine + 1
            rP = rP + 1
        End If
    Loop
    If np = 0 Then Exit Function
    ReDim Preserve parts(0 To np - 1)
    ReadString = Join(parts, vbNullString)
End Function

Private Sub AddPart(ByRef parts() As String, ByRef np As Long, ByVal s As String)
    If np > UBound(parts) Then
        ReDim Preserve parts(0 To 2 * np + 1)
    End If
    parts(np) = s
    np = np + 1
End Sub

Private Function ReadChar() As LispObj
    Dim st As Long, tok As String, ok As Boolean, c As String
    st = rP
    If rP > rN Then ReadErr "bad character literal"
    rP = rP + 1
    Do While rP <= rN
        If IsDelim(Mid$(rS, rP, 1)) Then Exit Do
        rP = rP + 1
    Loop
    tok = Mid$(rS, st, rP - st)
    If Len(tok) = 1 Then
        Set ReadChar = LMakeChar(tok)
        Exit Function
    End If
    Select Case LCase$(tok)
    Case "space": c = " "
    Case "newline", "linefeed", "nl": c = vbLf
    Case "tab": c = vbTab
    Case "return": c = vbCr
    Case "null", "nul": c = ChrW(0)
    Case "alarm": c = ChrW(7)
    Case "backspace": c = ChrW(8)
    Case "delete": c = ChrW(127)
    Case "escape": c = ChrW(27)
    Case Else
        If LCase$(Left$(tok, 1)) = "x" And Len(tok) > 1 Then
            c = ChrW(CLng(LParseInt(Mid$(tok, 2), 16, ok)))
            If Not ok Then ReadErr "bad character #\" & tok
        Else
            ' a single character followed by more symbol characters is invalid
            ReadErr "unknown character name #\" & tok
        End If
    End Select
    Set ReadChar = LMakeChar(c)
End Function
