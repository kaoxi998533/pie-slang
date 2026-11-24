Attribute VB_Name = "PyParser"
Option Explicit
' ===========================================================================
'  PyParser - tokenizer (with INDENT / DEDENT handling) and recursive-descent
'  parser that turns Python source text into a tree of PyNode objects.
' ===========================================================================

Private Const TT_NAME As Long = 1
Private Const TT_NUM As Long = 2
Private Const TT_STR As Long = 3
Private Const TT_FSTR As Long = 4
Private Const TT_OP As Long = 5
Private Const TT_NEWLINE As Long = 6
Private Const TT_INDENT As Long = 7
Private Const TT_DEDENT As Long = 8
Private Const TT_EOF As Long = 9

Private Const KEYWORDS As String = " False None True and as assert async await break class continue def del elif else except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield "

Private tkType() As Long
Private tkText() As String
Private tkVal() As Variant
Private tkLine() As Long
Private tkFlag() As Long
Private tkN As Long
Private P As Long
Private mCurDef As PyNode

' ---------------------------------------------------------------------------
'  public API
' ---------------------------------------------------------------------------
Public Function PyParse(ByVal source As String) As PyNode
    PrepTokens source
    Set mCurDef = Nothing
    Set PyParse = ParseFile()
End Function

Public Function PyParseExpr(ByVal source As String) As PyNode
    Dim e As PyNode
    PrepTokens source
    Set mCurDef = Nothing
    Do While tkType(P) = TT_NEWLINE
        P = P + 1
    Loop
    Set e = ParseTestListStar()
    Do While tkType(P) = TT_NEWLINE
        P = P + 1
    Loop
    If tkType(P) <> TT_EOF Then SynErr "invalid syntax"
    Set PyParseExpr = e
End Function

' Nothing when the source is not a single expression
Public Function PyTryParseExpr(ByVal source As String) As PyNode
    On Error GoTo Fail
    Set PyTryParseExpr = PyParseExpr(source)
    Exit Function
Fail:
    Set gExc = Nothing
    Set PyTryParseExpr = Nothing
End Function

Private Sub PrepTokens(ByVal source As String)
    Dim src As String
    src = Replace(source, vbCrLf, vbLf)
    src = Replace(src, vbCr, vbLf)
    If Len(src) > 0 Then
        If UCode(Left$(src, 1)) = 65279 Then src = Mid$(src, 2)
    End If
    ReDim tkType(0 To 255)
    ReDim tkText(0 To 255)
    ReDim tkVal(0 To 255)
    ReDim tkLine(0 To 255)
    ReDim tkFlag(0 To 255)
    tkN = 0
    Tokenize src, 1, False
    P = 0
End Sub

' ---------------------------------------------------------------------------
'  tokenizer
' ---------------------------------------------------------------------------
Private Sub AddTok(ByVal t As Long, ByVal txt As String, ByVal ln As Long)
    Dim nc As Long
    If tkN > UBound(tkType) Then
        nc = (UBound(tkType) + 1) * 2
        ReDim Preserve tkType(0 To nc - 1)
        ReDim Preserve tkText(0 To nc - 1)
        ReDim Preserve tkVal(0 To nc - 1)
        ReDim Preserve tkLine(0 To nc - 1)
        ReDim Preserve tkFlag(0 To nc - 1)
    End If
    tkType(tkN) = t
    tkText(tkN) = txt
    tkLine(tkN) = ln
    tkVal(tkN) = Empty
    tkFlag(tkN) = 0
    tkN = tkN + 1
End Sub

Private Sub LexErr(ByVal msg As String, ByVal ln As Long, ByVal indentErr As Boolean)
    gFrameLine(gDepth) = ln
    If indentErr Then
        ThrowErr "IndentationError", msg
    Else
        ThrowErr "SyntaxError", msg
    End If
End Sub

Private Function IsDigitC(ByVal c As String) As Boolean
    If Len(c) = 1 Then IsDigitC = (c >= "0" And c <= "9")
End Function

Private Function IsHexC(ByVal c As String) As Boolean
    If Len(c) = 1 Then IsHexC = (c >= "0" And c <= "9") Or (c >= "a" And c <= "f") Or (c >= "A" And c <= "F")
End Function

Private Function IsBadUnicode(ByVal k As Long) As Boolean
    ' punctuation blocks that must not be part of identifiers
    If k = 160 Then IsBadUnicode = True: Exit Function
    If k >= &H2000& And k <= &H206F& Then IsBadUnicode = True: Exit Function
    If k >= &H3000& And k <= &H303F& Then IsBadUnicode = True: Exit Function
    If k >= &HFF00& And k <= &HFF0F& Then IsBadUnicode = True: Exit Function
    If k >= &HFF1A& And k <= &HFF20& Then IsBadUnicode = True: Exit Function
    If k >= &HFF3B& And k <= &HFF40& Then IsBadUnicode = True: Exit Function
    If k >= &HFF5B& And k <= &HFF65& Then IsBadUnicode = True: Exit Function
    If k >= &HFE30& And k <= &HFE4F& Then IsBadUnicode = True: Exit Function
    If k = &HB7& Or k = &HD7& Or k = &HF7& Or k = &HAB& Or k = &HBB& Then IsBadUnicode = True
End Function

Private Function IsIdStart(ByVal c As String) As Boolean
    Dim k As Long
    If Len(c) <> 1 Then Exit Function
    k = UCode(c)
    If (k >= 65 And k <= 90) Or (k >= 97 And k <= 122) Or k = 95 Then
        IsIdStart = True
    ElseIf k >= 128 Then
        IsIdStart = Not IsBadUnicode(k)
    End If
End Function

Private Function IsIdChar(ByVal c As String) As Boolean
    If Len(c) <> 1 Then Exit Function
    If c >= "0" And c <= "9" Then
        IsIdChar = True
    Else
        IsIdChar = IsIdStart(c)
    End If
End Function

Private Function IsStrPrefix(ByVal w As String) As Boolean
    Select Case LCase$(w)
    Case "r", "u", "f", "b", "rb", "br", "fr", "rf"
        IsStrPrefix = True
    End Select
End Function

Private Sub Tokenize(ByRef src As String, ByVal firstLine As Long, ByVal exprMode As Boolean)
    Dim i As Long, n As Long, c As String, ln As Long, depth As Long, atBOL As Boolean, col As Long
    Dim j As Long, ind() As Long, nInd As Long, startTok As Long, word As String, q As String
    Dim two As String, three As String, brk As String, brkLn() As Long
    ReDim ind(0 To 63)
    ReDim brkLn(0 To 63)
    ind(0) = 0
    nInd = 1
    n = Len(src)
    i = 1
    ln = firstLine
    If exprMode Then depth = 1
    atBOL = Not exprMode
    startTok = tkN
    Do While i <= n
        If atBOL Then
            col = 0
            Do While i <= n
                c = Mid$(src, i, 1)
                If c = " " Or c = ChrW(160) Then
                    col = col + 1
                ElseIf c = vbTab Then
                    col = (col \ 8 + 1) * 8
                ElseIf c = vbFormFeed Then
                    col = 0
                Else
                    Exit Do
                End If
                i = i + 1
            Loop
            If i > n Then Exit Do
            c = Mid$(src, i, 1)
            If c = vbLf Then
                i = i + 1
                ln = ln + 1
                GoTo NextChar
            End If
            If c = "#" Then
                Do While i <= n
                    If Mid$(src, i, 1) = vbLf Then Exit Do
                    i = i + 1
                Loop
                GoTo NextChar
            End If
            If col > ind(nInd - 1) Then
                If nInd > UBound(ind) Then
                    ReDim Preserve ind(0 To 2 * nInd)
                End If
                ind(nInd) = col
                nInd = nInd + 1
                AddTok TT_INDENT, "", ln
            ElseIf col < ind(nInd - 1) Then
                Do While col < ind(nInd - 1)
                    nInd = nInd - 1
                    AddTok TT_DEDENT, "", ln
                Loop
                If col <> ind(nInd - 1) Then LexErr "unindent does not match any outer indentation level", ln, True
            End If
            atBOL = False
        End If
        c = Mid$(src, i, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbFormFeed Or c = ChrW(160) Then
            i = i + 1
        ElseIf c = vbLf Then
            i = i + 1
            If depth = 0 Then
                AddTok TT_NEWLINE, "", ln
                atBOL = True
            End If
            ln = ln + 1
        ElseIf c = "#" Then
            Do While i <= n
                If Mid$(src, i, 1) = vbLf Then Exit Do
                i = i + 1
            Loop
        ElseIf c = "\" Then
            If Mid$(src, i + 1, 1) = vbLf Then
                i = i + 2
                ln = ln + 1
            Else
                LexErr "unexpected character after line continuation character", ln, False
            End If
        ElseIf c = """" Or c = "'" Then
            LexString src, i, ln, ""
        ElseIf IsIdStart(c) Then
            j = i + 1
            Do While j <= n
                If Not IsIdChar(Mid$(src, j, 1)) Then Exit Do
                j = j + 1
            Loop
            word = Mid$(src, i, j - i)
            q = Mid$(src, j, 1)
            If (q = """" Or q = "'") And IsStrPrefix(word) Then
                i = j
                LexString src, i, ln, LCase$(word)
            Else
                AddTok TT_NAME, word, ln
                i = j
            End If
        ElseIf IsDigitC(c) Then
            LexNumber src, i, ln
        ElseIf c = "." And IsDigitC(Mid$(src, i + 1, 1)) Then
            LexNumber src, i, ln
        Else
            three = Mid$(src, i, 3)
            two = Mid$(src, i, 2)
            If three = "**=" Or three = "//=" Or three = ">>=" Or three = "<<=" Or three = "..." Then
                AddTok TT_OP, three, ln
                i = i + 3
            ElseIf InStr(1, " ** // == != <= >= += -= *= /= %= &= |= ^= -> << >> := @= ", " " & two & " ") > 0 Then
                AddTok TT_OP, two, ln
                i = i + 2
            ElseIf InStr(1, "+-*/%@&|^~<>()[]{},:.;=!", c) > 0 Then
                If c = "(" Or c = "[" Or c = "{" Then
                    depth = depth + 1
                    If Len(brk) <= UBound(brkLn) Then brkLn(Len(brk)) = ln
                    brk = brk & c
                ElseIf c = ")" Or c = "]" Or c = "}" Then
                    If depth > 0 Then depth = depth - 1
                    If Len(brk) > 0 Then
                        If InStr(1, "([{", Right$(brk, 1)) <> InStr(1, ")]}", c) Then
                            LexErr "closing parenthesis '" & c & "' does not match opening parenthesis '" & Right$(brk, 1) & "'", ln, False
                        End If
                        brk = Left$(brk, Len(brk) - 1)
                    ElseIf Not exprMode Then
                        LexErr "unmatched '" & c & "'", ln, False
                    End If
                End If
                AddTok TT_OP, c, ln
                i = i + 1
            Else
                LexErr "invalid character '" & c & "' (U+" & Right$("000" & Hex$(UCode(c)), 4) & ")", ln, False
            End If
        End If
NextChar:
    Loop
    If Not exprMode And Len(brk) > 0 Then
        If Len(brk) - 1 <= UBound(brkLn) Then
            LexErr "'" & Right$(brk, 1) & "' was never closed", brkLn(Len(brk) - 1), False
        Else
            LexErr "'" & Right$(brk, 1) & "' was never closed", ln, False
        End If
    End If
    If Not exprMode Then
        If tkN > startTok Then
            If tkType(tkN - 1) <> TT_NEWLINE Then AddTok TT_NEWLINE, "", ln
        End If
        Do While nInd > 1
            nInd = nInd - 1
            AddTok TT_DEDENT, "", ln
        Loop
    End If
    AddTok TT_EOF, "", ln
End Sub

Private Sub LexString(ByRef src As String, ByRef i As Long, ByRef ln As Long, ByVal prefix As String)
    Dim q As String, triple As Boolean, raw As Boolean, isF As Boolean, n As Long, c As String
    Dim startPos As Long, endPos As Long, startLn As Long, rawText As String
    n = Len(src)
    raw = (InStr(1, prefix, "r") > 0)
    isF = (InStr(1, prefix, "f") > 0)
    q = Mid$(src, i, 1)
    startLn = ln
    If Mid$(src, i, 3) = q & q & q Then
        triple = True
        i = i + 3
    Else
        i = i + 1
    End If
    startPos = i
    Do
        If i > n Then LexErr "unterminated string literal", startLn, False
        c = Mid$(src, i, 1)
        If c = "\" Then
            If Mid$(src, i + 1, 1) = vbLf Then ln = ln + 1
            i = i + 2
        ElseIf c = q Then
            If Not triple Then
                endPos = i
                i = i + 1
                Exit Do
            End If
            If Mid$(src, i, 3) = q & q & q Then
                endPos = i
                i = i + 3
                Exit Do
            End If
            i = i + 1
        ElseIf c = vbLf Then
            If Not triple Then LexErr "unterminated string literal", startLn, False
            ln = ln + 1
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
    rawText = Mid$(src, startPos, endPos - startPos)
    If isF Then
        AddTok TT_FSTR, rawText, startLn
        If raw Then tkFlag(tkN - 1) = 1
    ElseIf raw Then
        AddTok TT_STR, rawText, startLn
    Else
        AddTok TT_STR, Unescape(rawText, startLn), startLn
    End If
End Sub

Private Sub AddPart(ByRef parts() As String, ByRef np As Long, ByVal s As String)
    If np > UBound(parts) Then
        ReDim Preserve parts(0 To 2 * np + 1)
    End If
    parts(np) = s
    np = np + 1
End Sub

Private Function HexVal(ByVal hx As String, ByVal want As Long, ByVal ln As Long) As Long
    Dim i As Long, c As String, v As Long, d As Long
    If Len(hx) <> want Then LexErr "truncated \x / \u escape", ln, False
    For i = 1 To Len(hx)
        c = LCase$(Mid$(hx, i, 1))
        If c >= "0" And c <= "9" Then
            d = Asc(c) - 48
        ElseIf c >= "a" And c <= "f" Then
            d = Asc(c) - 87
        Else
            LexErr "invalid hex escape", ln, False
        End If
        v = v * 16 + d
    Next
    HexVal = v
End Function

Private Function CodeToStr(ByVal code As Long) As String
    If code > 65535 Then
        code = code - 65536
        CodeToStr = ChrW(&HD800& + (code \ 1024)) & ChrW(&HDC00& + (code And 1023))
    Else
        CodeToStr = ChrW(code)
    End If
End Function

Private Function Unescape(ByVal s As String, ByVal ln As Long) As String
    Dim p As Long, st As Long, parts() As String, np As Long, c As String, n As Long, k As Long, v As Long
    p = InStr(1, s, "\")
    If p = 0 Then
        Unescape = s
        Exit Function
    End If
    n = Len(s)
    ReDim parts(0 To 15)
    st = 1
    Do While p > 0
        AddPart parts, np, Mid$(s, st, p - st)
        c = Mid$(s, p + 1, 1)
        st = p + 2
        Select Case c
        Case "n": AddPart parts, np, vbLf
        Case "t": AddPart parts, np, vbTab
        Case "r": AddPart parts, np, vbCr
        Case "\": AddPart parts, np, "\"
        Case "'": AddPart parts, np, "'"
        Case """": AddPart parts, np, """"
        Case "a": AddPart parts, np, ChrW(7)
        Case "b": AddPart parts, np, ChrW(8)
        Case "f": AddPart parts, np, ChrW(12)
        Case "v": AddPart parts, np, ChrW(11)
        Case "0", "1", "2", "3", "4", "5", "6", "7"
            v = 0
            k = p + 1
            Do While k <= n And k <= p + 3
                c = Mid$(s, k, 1)
                If c < "0" Or c > "7" Then Exit Do
                v = v * 8 + (Asc(c) - 48)
                k = k + 1
            Loop
            AddPart parts, np, ChrW(v)
            st = k
        Case "x"
            AddPart parts, np, ChrW(HexVal(Mid$(s, p + 2, 2), 2, ln))
            st = p + 4
        Case "u"
            AddPart parts, np, CodeToStr(HexVal(Mid$(s, p + 2, 4), 4, ln))
            st = p + 6
        Case "U"
            AddPart parts, np, CodeToStr(HexVal(Mid$(s, p + 2, 8), 8, ln))
            st = p + 10
        Case vbLf
            ' line continuation inside a string
        Case Else
            AddPart parts, np, "\" & c
        End Select
        If st > n Then
            p = 0
        Else
            p = InStr(st, s, "\")
        End If
    Loop
    If st <= n Then AddPart parts, np, Mid$(s, st)
    ReDim Preserve parts(0 To np - 1)
    Unescape = Join(parts, vbNullString)
End Function

Private Sub LexNumber(ByRef src As String, ByRef i As Long, ByVal ln As Long)
    Dim j As Long, n As Long, s As String, c As String, c2 As String, isFloat As Boolean, hasExp As Boolean
    Dim pref As String, bs As Long, ok As Boolean, v As Variant
    n = Len(src)
    pref = LCase$(Mid$(src, i, 2))
    If pref = "0x" Or pref = "0b" Or pref = "0o" Then
        j = i + 2
        Do While j <= n
            c = Mid$(src, j, 1)
            If Not (IsHexC(c) Or c = "_") Then Exit Do
            j = j + 1
        Loop
        Select Case pref
        Case "0x": bs = 16
        Case "0b": bs = 2
        Case Else: bs = 8
        End Select
        v = ParseIntStr(Mid$(src, i + 2, j - i - 2), bs, ok)
        If Not ok Then LexErr "invalid number literal", ln, False
        AddTok TT_NUM, Mid$(src, i, j - i), ln
        tkVal(tkN - 1) = v
        i = j
        Exit Sub
    End If
    j = i
    Do While j <= n
        c = Mid$(src, j, 1)
        If IsDigitC(c) Or c = "_" Then
            j = j + 1
        ElseIf c = "." And Not isFloat Then
            isFloat = True
            j = j + 1
        ElseIf (c = "e" Or c = "E") And Not hasExp Then
            c2 = Mid$(src, j + 1, 1)
            If IsDigitC(c2) Then
                hasExp = True
                isFloat = True
                j = j + 2
            ElseIf (c2 = "+" Or c2 = "-") And IsDigitC(Mid$(src, j + 2, 1)) Then
                hasExp = True
                isFloat = True
                j = j + 3
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    s = Replace(Mid$(src, i, j - i), "_", vbNullString)
    c = Mid$(src, j, 1)
    If c = "j" Or c = "J" Then LexErr "complex numbers are not supported", ln, False
    If IsIdStart(c) Then LexErr "invalid decimal literal", ln, False
    If isFloat Then
        v = CDbl(Val(s))
    Else
        v = ParseIntStr(s, 10, ok)
        If Not ok Then LexErr "invalid number literal", ln, False
    End If
    AddTok TT_NUM, s, ln
    tkVal(tkN - 1) = v
    i = j
End Sub

' ---------------------------------------------------------------------------
'  parser helpers
' ---------------------------------------------------------------------------
Private Function NewNode(ByVal kind As Long, ByVal ln As Long) As PyNode
    Dim nd As PyNode
    Set nd = New PyNode
    nd.Kind = kind
    nd.Ln = ln
    Set NewNode = nd
End Function

Private Function IsReserved(ByVal w As String) As Boolean
    IsReserved = (InStr(1, KEYWORDS, " " & w & " ", vbBinaryCompare) > 0)
End Function

Private Function IsOp(ByVal s As String) As Boolean
    If tkType(P) = TT_OP Then IsOp = (tkText(P) = s)
End Function

Private Function IsKw(ByVal s As String) As Boolean
    If tkType(P) = TT_NAME Then IsKw = (tkText(P) = s)
End Function

Private Function AcceptOp(ByVal s As String) As Boolean
    If IsOp(s) Then
        P = P + 1
        AcceptOp = True
    End If
End Function

Private Function AcceptKw(ByVal s As String) As Boolean
    If IsKw(s) Then
        P = P + 1
        AcceptKw = True
    End If
End Function

Private Sub ExpectOp(ByVal s As String)
    If Not AcceptOp(s) Then SynErr "expected '" & s & "'"
End Sub

Private Sub ExpectKw(ByVal s As String)
    If Not AcceptKw(s) Then SynErr "expected '" & s & "'"
End Sub

Private Function PeekOpAt(ByVal k As Long, ByVal s As String) As Boolean
    If P + k < tkN Then
        If tkType(P + k) = TT_OP Then PeekOpAt = (tkText(P + k) = s)
    End If
End Function

Private Function TokDesc() As String
    Select Case tkType(P)
    Case TT_NEWLINE: TokDesc = " (at end of line)"
    Case TT_EOF: TokDesc = " (unexpected end of input)"
    Case TT_INDENT: TokDesc = " (unexpected indent)"
    Case TT_DEDENT: TokDesc = " (unexpected dedent)"
    Case TT_STR, TT_FSTR: TokDesc = " (near a string literal)"
    Case Else: TokDesc = " (near '" & tkText(P) & "')"
    End Select
End Function

Private Sub SynErr(ByVal msg As String)
    Dim ln As Long
    If P < tkN Then ln = tkLine(P)
    LexErr msg & TokDesc(), ln, False
End Sub

Private Function ExpectName() As String
    If tkType(P) <> TT_NAME Then SynErr "expected a name"
    If IsReserved(tkText(P)) Then SynErr "invalid syntax"
    ExpectName = tkText(P)
    P = P + 1
End Function

Private Function NameNode(ByVal nm As String, ByVal ln As Long) As PyNode
    Dim nd As PyNode
    Set nd = NewNode(NK_NAME, ln)
    nd.S = nm
    nd.H = HashStr(nm)
    Set NameNode = nd
End Function

Private Function AtStmtEnd() As Boolean
    Select Case tkType(P)
    Case TT_NEWLINE, TT_EOF, TT_DEDENT
        AtStmtEnd = True
    Case TT_OP
        AtStmtEnd = (tkText(P) = ";")
    End Select
End Function

Private Function AtExprEnd() As Boolean
    Select Case tkType(P)
    Case TT_NEWLINE, TT_EOF, TT_DEDENT, TT_INDENT
        AtExprEnd = True
    Case TT_OP
        Select Case tkText(P)
        Case ")", "]", "}", "=", ":", ";", "+=", "-=", "*=", "/=", "//=", "%=", "**=", "&=", "|=", "^=", ">>=", "<<=", "@="
            AtExprEnd = True
        End Select
    Case TT_NAME
        AtExprEnd = (tkText(P) = "in")
    End Select
End Function

' ---------------------------------------------------------------------------
'  statements
' ---------------------------------------------------------------------------
Private Function ParseFile() As PyNode
    Dim blk As PyNode
    Set blk = NewNode(NK_BLOCK, 1)
    Do While tkType(P) <> TT_EOF
        If tkType(P) = TT_NEWLINE Then
            P = P + 1
        Else
            ParseStatement blk
        End If
    Loop
    Set ParseFile = blk
End Function

Private Sub ParseStatement(ByVal blk As PyNode)
    If tkType(P) = TT_INDENT Then LexErr "unexpected indent", tkLine(P), True
    If tkType(P) = TT_NAME Then
        Select Case tkText(P)
        Case "if"
            blk.Add ParseIf()
            Exit Sub
        Case "while"
            blk.Add ParseWhile()
            Exit Sub
        Case "for"
            blk.Add ParseFor()
            Exit Sub
        Case "def"
            blk.Add ParseDef(Nothing)
            Exit Sub
        Case "class"
            blk.Add ParseClass(Nothing)
            Exit Sub
        Case "try"
            blk.Add ParseTry()
            Exit Sub
        Case "with"
            SynErr "the 'with' statement is not supported"
        Case "async"
            SynErr "async code is not supported"
        End Select
    ElseIf IsOp("@") Then
        blk.Add ParseDecorated()
        Exit Sub
    End If
    ParseSimpleStmts blk
End Sub

Private Sub ParseSimpleStmts(ByVal blk As PyNode)
    Do
        blk.Add ParseSmallStmt()
        If Not AcceptOp(";") Then Exit Do
        If tkType(P) = TT_NEWLINE Or tkType(P) = TT_EOF Then Exit Do
    Loop
    If tkType(P) = TT_NEWLINE Then
        P = P + 1
    ElseIf tkType(P) <> TT_EOF And tkType(P) <> TT_DEDENT Then
        SynErr "invalid syntax"
    End If
End Sub

Private Function ParseBlock() As PyNode
    Dim blk As PyNode
    Set blk = NewNode(NK_BLOCK, tkLine(P))
    If tkType(P) = TT_NEWLINE Then
        P = P + 1
        If tkType(P) <> TT_INDENT Then LexErr "expected an indented block", tkLine(P), True
        P = P + 1
        Do While tkType(P) <> TT_DEDENT And tkType(P) <> TT_EOF
            If tkType(P) = TT_NEWLINE Then
                P = P + 1
            Else
                ParseStatement blk
            End If
        Loop
        If tkType(P) = TT_DEDENT Then P = P + 1
    Else
        ParseSimpleStmts blk
    End If
    Set ParseBlock = blk
End Function

Private Sub CheckTarget(ByVal t As PyNode, ByVal allowStar As Boolean)
    Dim i As Long, nStar As Long
    Select Case t.Kind
    Case NK_NAME, NK_ATTR, NK_SUBSCR
    Case NK_TUPLE, NK_LIST
        For i = 0 To t.N - 1
            If t.Kid(i).Kind = NK_STARRED Then
                nStar = nStar + 1
                CheckTarget t.Kid(i).A, False
            Else
                CheckTarget t.Kid(i), True
            End If
        Next
        If nStar > 1 Then LexErr "multiple starred expressions in assignment", t.Ln, False
    Case NK_CALL
        LexErr "cannot assign to function call", t.Ln, False
    Case NK_CONST
        LexErr "cannot assign to literal", t.Ln, False
    Case Else
        LexErr "cannot assign to expression", t.Ln, False
    End Select
End Sub

Private Function AugOpCode() As Long
    If tkType(P) <> TT_OP Then Exit Function
    Select Case tkText(P)
    Case "+=": AugOpCode = OP_ADD
    Case "-=": AugOpCode = OP_SUB
    Case "*=": AugOpCode = OP_MUL
    Case "/=": AugOpCode = OP_DIV
    Case "//=": AugOpCode = OP_FLOORDIV
    Case "%=": AugOpCode = OP_MOD
    Case "**=": AugOpCode = OP_POW
    Case "&=": AugOpCode = OP_BITAND
    Case "|=": AugOpCode = OP_BITOR
    Case "^=": AugOpCode = OP_BITXOR
    Case ">>=": AugOpCode = OP_RSHIFT
    Case "<<=": AugOpCode = OP_LSHIFT
    Case "@=": AugOpCode = OP_MATMUL
    End Select
End Function

Private Function ParseDottedName() As String
    Dim s As String
    s = ExpectName()
    Do While AcceptOp(".")
        s = s & "." & ExpectName()
    Loop
    ParseDottedName = s
End Function

Private Function ParseSmallStmt() As PyNode
    Dim nd As PyNode, ln As Long, t As String, e As PyNode, k As PyNode, i As Long, op As Long
    Dim exprs As PyNode, paren As Boolean
    ln = tkLine(P)
    If tkType(P) = TT_NAME Then
        t = tkText(P)
        Select Case t
        Case "pass"
            P = P + 1
            Set ParseSmallStmt = NewNode(NK_PASS, ln)
            Exit Function
        Case "break"
            P = P + 1
            Set ParseSmallStmt = NewNode(NK_BREAK, ln)
            Exit Function
        Case "continue"
            P = P + 1
            Set ParseSmallStmt = NewNode(NK_CONTINUE, ln)
            Exit Function
        Case "return"
            P = P + 1
            Set nd = NewNode(NK_RETURN, ln)
            If Not AtStmtEnd() Then Set nd.A = ParseTestListStar()
            Set ParseSmallStmt = nd
            Exit Function
        Case "global", "nonlocal"
            P = P + 1
            If t = "global" Then Set nd = NewNode(NK_GLOBAL, ln) Else Set nd = NewNode(NK_NONLOCAL, ln)
            Do
                nd.Add NameNode(ExpectName(), ln)
                If Not AcceptOp(",") Then Exit Do
            Loop
            Set ParseSmallStmt = nd
            Exit Function
        Case "del"
            P = P + 1
            Set nd = NewNode(NK_DEL, ln)
            Set e = ParseTargetList()
            If e.Kind = NK_TUPLE Then
                For i = 0 To e.N - 1
                    nd.Add e.Kid(i)
                Next
            Else
                nd.Add e
            End If
            Set ParseSmallStmt = nd
            Exit Function
        Case "yield"
            P = P + 1
            If mCurDef Is Nothing Then LexErr "'yield' outside function", ln, False
            mCurDef.Op = mCurDef.Op Or 2
            Set nd = NewNode(NK_YIELD, ln)
            If AcceptKw("from") Then
                nd.Op = 1
                Set nd.A = ParseTest()
            ElseIf Not AtStmtEnd() Then
                Set nd.A = ParseTestListStar()
            End If
            If IsOp("=") Then LexErr "yield expressions are only supported as statements", ln, False
            Set ParseSmallStmt = nd
            Exit Function
        Case "assert"
            P = P + 1
            Set nd = NewNode(NK_ASSERT, ln)
            Set nd.A = ParseTest()
            If AcceptOp(",") Then Set nd.B = ParseTest()
            Set ParseSmallStmt = nd
            Exit Function
        Case "raise"
            P = P + 1
            Set nd = NewNode(NK_RAISE, ln)
            If Not AtStmtEnd() Then
                Set nd.A = ParseTest()
                If AcceptKw("from") Then Set nd.B = ParseTest()
            End If
            Set ParseSmallStmt = nd
            Exit Function
        Case "import"
            P = P + 1
            Set nd = NewNode(NK_IMPORT, ln)
            Do
                Set k = NewNode(NK_NAME, ln)
                k.S = ParseDottedName()
                If AcceptKw("as") Then k.V = ExpectName()
                nd.Add k
                If Not AcceptOp(",") Then Exit Do
            Loop
            Set ParseSmallStmt = nd
            Exit Function
        Case "from"
            P = P + 1
            Set nd = NewNode(NK_FROMIMPORT, ln)
            nd.S = ParseDottedName()
            ExpectKw "import"
            If AcceptOp("*") Then
                Set k = NewNode(NK_NAME, ln)
                k.S = "*"
                nd.Add k
            Else
                paren = AcceptOp("(")
                Do
                    Set k = NewNode(NK_NAME, ln)
                    k.S = ExpectName()
                    If AcceptKw("as") Then k.V = ExpectName()
                    nd.Add k
                    If Not AcceptOp(",") Then Exit Do
                    If paren And IsOp(")") Then Exit Do
                Loop
                If paren Then ExpectOp ")"
            End If
            Set ParseSmallStmt = nd
            Exit Function
        End Select
    End If
    ' expression statement or assignment
    Set e = ParseTestListStar()
    If IsOp("=") Then
        Set exprs = NewNode(NK_BLOCK, ln)
        exprs.Add e
        Do While AcceptOp("=")
            exprs.Add ParseTestListStar()
        Loop
        Set nd = NewNode(NK_ASSIGN, ln)
        Set nd.A = exprs.Kid(exprs.N - 1)
        For i = 0 To exprs.N - 2
            CheckTarget exprs.Kid(i), True
            nd.Add exprs.Kid(i)
        Next
        Set ParseSmallStmt = nd
        Exit Function
    End If
    op = AugOpCode()
    If op <> 0 Then
        If e.Kind <> NK_NAME And e.Kind <> NK_ATTR And e.Kind <> NK_SUBSCR Then
            LexErr "illegal expression for augmented assignment", ln, False
        End If
        P = P + 1
        Set nd = NewNode(NK_AUGASSIGN, ln)
        nd.Op = op
        Set nd.A = e
        Set nd.B = ParseTestListStar()
        Set ParseSmallStmt = nd
        Exit Function
    End If
    If IsOp(":") Then
        ' annotated assignment: the annotation itself is ignored
        P = P + 1
        ParseTest
        CheckTarget e, False
        Set nd = NewNode(NK_ANNASSIGN, ln)
        Set nd.A = e
        If AcceptOp("=") Then Set nd.B = ParseTestListStar()
        Set ParseSmallStmt = nd
        Exit Function
    End If
    Set nd = NewNode(NK_EXPR, ln)
    Set nd.A = e
    Set ParseSmallStmt = nd
End Function

Private Function ParseIf() As PyNode
    Dim nd As PyNode, ln As Long
    ln = tkLine(P)
    P = P + 1
    Set nd = NewNode(NK_IF, ln)
    Set nd.A = ParseTest()
    ExpectOp ":"
    Set nd.B = ParseBlock()
    If IsKw("elif") Then
        Set nd.C = ParseIf()
    ElseIf IsKw("else") Then
        P = P + 1
        ExpectOp ":"
        Set nd.C = ParseBlock()
    End If
    Set ParseIf = nd
End Function

Private Function ParseWhile() As PyNode
    Dim nd As PyNode, ln As Long
    ln = tkLine(P)
    P = P + 1
    Set nd = NewNode(NK_WHILE, ln)
    Set nd.A = ParseTest()
    ExpectOp ":"
    Set nd.B = ParseBlock()
    If IsKw("else") Then
        P = P + 1
        ExpectOp ":"
        Set nd.C = ParseBlock()
    End If
    Set ParseWhile = nd
End Function

Private Function ParseFor() As PyNode
    Dim nd As PyNode, ln As Long
    ln = tkLine(P)
    P = P + 1
    Set nd = NewNode(NK_FOR, ln)
    Set nd.A = ParseTargetList()
    CheckTarget nd.A, True
    ExpectKw "in"
    Set nd.B = ParseTestListStar()
    ExpectOp ":"
    Set nd.C = ParseBlock()
    If IsKw("else") Then
        P = P + 1
        ExpectOp ":"
        Set nd.D = ParseBlock()
    End If
    Set ParseFor = nd
End Function

Private Function ParseTry() As PyNode
    Dim nd As PyNode, ln As Long, h As PyNode
    ln = tkLine(P)
    P = P + 1
    Set nd = NewNode(NK_TRY, ln)
    ExpectOp ":"
    Set nd.B = ParseBlock()
    Do While IsKw("except")
        Set h = NewNode(NK_HANDLER, tkLine(P))
        P = P + 1
        If Not IsOp(":") Then
            Set h.A = ParseTest()
            If AcceptKw("as") Then
                h.S = ExpectName()
                h.H = HashStr(h.S)
            End If
        End If
        ExpectOp ":"
        Set h.B = ParseBlock()
        nd.Add h
    Loop
    If IsKw("else") Then
        If nd.N = 0 Then SynErr "expected 'except' or 'finally' block"
        P = P + 1
        ExpectOp ":"
        Set nd.C = ParseBlock()
    End If
    If IsKw("finally") Then
        P = P + 1
        ExpectOp ":"
        Set nd.D = ParseBlock()
    End If
    If nd.N = 0 And nd.D Is Nothing Then SynErr "expected 'except' or 'finally' block"
    Set ParseTry = nd
End Function

Private Function ParseDecorated() As PyNode
    Dim decos As PyNode
    Set decos = NewNode(NK_BLOCK, tkLine(P))
    Do While AcceptOp("@")
        decos.Add ParseTest()
        If tkType(P) <> TT_NEWLINE Then SynErr "invalid syntax"
        P = P + 1
    Loop
    If IsKw("def") Then
        Set ParseDecorated = ParseDef(decos)
    ElseIf IsKw("class") Then
        Set ParseDecorated = ParseClass(decos)
    Else
        SynErr "expected 'def' or 'class' after a decorator"
    End If
End Function

Private Sub ParseParams(ByVal nd As PyNode, ByVal closer As String)
    Dim pn As PyNode, kwOnly As Boolean
    Do While Not IsOp(closer)
        Set pn = NewNode(NK_PARAM, tkLine(P))
        If AcceptOp("**") Then
            pn.Op = 2
            pn.S = ExpectName()
        ElseIf AcceptOp("*") Then
            If IsOp(",") Or IsOp(closer) Then
                kwOnly = True
                GoTo NextParam
            End If
            pn.Op = 1
            pn.S = ExpectName()
            kwOnly = True
        ElseIf AcceptOp("/") Then
            GoTo NextParam
        Else
            pn.S = ExpectName()
            If kwOnly Then
                pn.Op = 3
            Else
                pn.Op = 0
            End If
        End If
        pn.H = HashStr(pn.S)
        If closer = ")" Then
            If AcceptOp(":") Then ParseTest
        End If
        If AcceptOp("=") Then Set pn.A = ParseTest()
        nd.Add pn
NextParam:
        If Not AcceptOp(",") Then Exit Do
    Loop
End Sub

Private Function ParseDef(ByVal decos As PyNode) As PyNode
    Dim nd As PyNode, ln As Long, saveDef As PyNode
    ln = tkLine(P)
    P = P + 1
    Set nd = NewNode(NK_DEF, ln)
    nd.S = ExpectName()
    nd.H = HashStr(nd.S)
    ExpectOp "("
    ParseParams nd, ")"
    ExpectOp ")"
    If AcceptOp("->") Then ParseTest
    ExpectOp ":"
    Set saveDef = mCurDef
    Set mCurDef = nd
    Set nd.B = ParseBlock()
    Set mCurDef = saveDef
    Set nd.C = decos
    Set ParseDef = nd
End Function

Private Function ParseClass(ByVal decos As PyNode) As PyNode
    Dim nd As PyNode, ln As Long, saveDef As PyNode
    ln = tkLine(P)
    P = P + 1
    Set nd = NewNode(NK_CLASS, ln)
    nd.S = ExpectName()
    nd.H = HashStr(nd.S)
    If AcceptOp("(") Then
        Do While Not IsOp(")")
            If tkType(P) = TT_NAME And PeekOpAt(1, "=") Then
                P = P + 2
                ParseTest
            Else
                nd.Add ParseTest()
            End If
            If Not AcceptOp(",") Then Exit Do
        Loop
        ExpectOp ")"
    End If
    ExpectOp ":"
    Set saveDef = mCurDef
    Set mCurDef = Nothing
    Set nd.B = ParseBlock()
    Set mCurDef = saveDef
    Set nd.C = decos
    Set ParseClass = nd
End Function

' ---------------------------------------------------------------------------
'  expressions
' ---------------------------------------------------------------------------
Private Function ParseTestListStar() As PyNode
    Dim first As PyNode, tup As PyNode, ln As Long
    ln = tkLine(P)
    Set first = ParseTestOrStar()
    If Not IsOp(",") Then
        Set ParseTestListStar = first
        Exit Function
    End If
    Set tup = NewNode(NK_TUPLE, ln)
    tup.Add first
    Do While AcceptOp(",")
        If AtExprEnd() Then Exit Do
        tup.Add ParseTestOrStar()
    Loop
    Set ParseTestListStar = tup
End Function

Private Function ParseTestOrStar() As PyNode
    Dim nd As PyNode
    If IsOp("*") Then
        Set nd = NewNode(NK_STARRED, tkLine(P))
        P = P + 1
        Set nd.A = ParseBin(1)
        Set ParseTestOrStar = nd
    Else
        Set ParseTestOrStar = ParseTest()
    End If
End Function

Private Function ParseTargetItem() As PyNode
    Dim nd As PyNode
    If IsOp("*") Then
        Set nd = NewNode(NK_STARRED, tkLine(P))
        P = P + 1
        Set nd.A = ParseBin(1)
        Set ParseTargetItem = nd
    Else
        Set ParseTargetItem = ParseBin(1)
    End If
End Function

Private Function ParseTargetList() As PyNode
    Dim first As PyNode, tup As PyNode
    Set first = ParseTargetItem()
    If Not IsOp(",") Then
        Set ParseTargetList = first
        Exit Function
    End If
    Set tup = NewNode(NK_TUPLE, first.Ln)
    tup.Add first
    Do While AcceptOp(",")
        If AtExprEnd() Then Exit Do
        tup.Add ParseTargetItem()
    Loop
    Set ParseTargetList = tup
End Function

Private Function ParseTest() As PyNode
    Dim c As PyNode, nd As PyNode, ln As Long
    If IsKw("lambda") Then
        Set ParseTest = ParseLambda()
        Exit Function
    End If
    ln = tkLine(P)
    If tkType(P) = TT_NAME And PeekOpAt(1, ":=") Then
        Set nd = NewNode(NK_WALRUS, ln)
        nd.S = ExpectName()
        nd.H = HashStr(nd.S)
        P = P + 1
        Set nd.A = ParseTest()
        Set ParseTest = nd
        Exit Function
    End If
    Set c = ParseOrTest()
    If IsKw("if") Then
        P = P + 1
        Set nd = NewNode(NK_IFEXP, ln)
        Set nd.B = c
        Set nd.A = ParseOrTest()
        ExpectKw "else"
        Set nd.C = ParseTest()
        Set ParseTest = nd
    Else
        Set ParseTest = c
    End If
End Function

Private Function ParseLambda() As PyNode
    Dim nd As PyNode
    Set nd = NewNode(NK_LAMBDA, tkLine(P))
    P = P + 1
    nd.S = "<lambda>"
    ParseParams nd, ":"
    ExpectOp ":"
    Set nd.A = ParseTest()
    Set ParseLambda = nd
End Function

Private Function ParseOrTest() As PyNode
    Dim l As PyNode, nd As PyNode
    Set l = ParseAndTest()
    Do While IsKw("or")
        Set nd = NewNode(NK_OR, tkLine(P))
        P = P + 1
        Set nd.A = l
        Set nd.B = ParseAndTest()
        Set l = nd
    Loop
    Set ParseOrTest = l
End Function

Private Function ParseAndTest() As PyNode
    Dim l As PyNode, nd As PyNode
    Set l = ParseNotTest()
    Do While IsKw("and")
        Set nd = NewNode(NK_AND, tkLine(P))
        P = P + 1
        Set nd.A = l
        Set nd.B = ParseNotTest()
        Set l = nd
    Loop
    Set ParseAndTest = l
End Function

Private Function ParseNotTest() As PyNode
    Dim nd As PyNode
    If IsKw("not") Then
        Set nd = NewNode(NK_NOT, tkLine(P))
        P = P + 1
        Set nd.A = ParseNotTest()
        Set ParseNotTest = nd
    Else
        Set ParseNotTest = ParseComparison()
    End If
End Function

Private Function CompOpAt(ByVal consume As Boolean) As Long
    Dim op As Long, adv As Long
    If tkType(P) = TT_OP Then
        Select Case tkText(P)
        Case "<": op = OP_LT
        Case ">": op = OP_GT
        Case "==": op = OP_EQ
        Case ">=": op = OP_GE
        Case "<=": op = OP_LE
        Case "!=": op = OP_NE
        End Select
        adv = 1
    ElseIf tkType(P) = TT_NAME Then
        Select Case tkText(P)
        Case "in"
            op = OP_IN
            adv = 1
        Case "not"
            If tkType(P + 1) = TT_NAME Then
                If tkText(P + 1) = "in" Then
                    op = OP_NOTIN
                    adv = 2
                End If
            End If
        Case "is"
            op = OP_IS
            adv = 1
            If tkType(P + 1) = TT_NAME Then
                If tkText(P + 1) = "not" Then
                    op = OP_ISNOT
                    adv = 2
                End If
            End If
        End Select
    End If
    If consume And op <> 0 Then P = P + adv
    CompOpAt = op
End Function

Private Function ParseComparison() As PyNode
    Dim first As PyNode, nd As PyNode, op As Long, ops As String, nxt As PyNode, ln As Long
    ln = tkLine(P)
    Set first = ParseBin(1)
    op = CompOpAt(True)
    If op = 0 Then
        Set ParseComparison = first
        Exit Function
    End If
    Set nxt = ParseBin(1)
    If CompOpAt(False) = 0 Then
        Set nd = NewNode(NK_COMPARE, ln)
        nd.Op = op
        Set nd.A = first
        Set nd.B = nxt
        Set ParseComparison = nd
        Exit Function
    End If
    Set nd = NewNode(NK_CHAIN, ln)
    nd.Add first
    nd.Add nxt
    ops = Chr$(op)
    Do
        op = CompOpAt(True)
        If op = 0 Then Exit Do
        ops = ops & Chr$(op)
        nd.Add ParseBin(1)
    Loop
    nd.S = ops
    Set ParseComparison = nd
End Function

Private Function BinOpAt(ByVal lvl As Long) As Long
    If tkType(P) <> TT_OP Then Exit Function
    Select Case lvl
    Case 1
        If tkText(P) = "|" Then BinOpAt = OP_BITOR
    Case 2
        If tkText(P) = "^" Then BinOpAt = OP_BITXOR
    Case 3
        If tkText(P) = "&" Then BinOpAt = OP_BITAND
    Case 4
        Select Case tkText(P)
        Case "<<": BinOpAt = OP_LSHIFT
        Case ">>": BinOpAt = OP_RSHIFT
        End Select
    Case 5
        Select Case tkText(P)
        Case "+": BinOpAt = OP_ADD
        Case "-": BinOpAt = OP_SUB
        End Select
    Case 6
        Select Case tkText(P)
        Case "*": BinOpAt = OP_MUL
        Case "/": BinOpAt = OP_DIV
        Case "//": BinOpAt = OP_FLOORDIV
        Case "%": BinOpAt = OP_MOD
        Case "@": BinOpAt = OP_MATMUL
        End Select
    End Select
End Function

' binary operator levels: 1 |   2 ^   3 &   4 << >>   5 + -   6 * / // % @
Private Function ParseBin(ByVal lvl As Long) As PyNode
    Dim l As PyNode, nd As PyNode, op As Long
    If lvl > 6 Then
        Set ParseBin = ParseFactor()
        Exit Function
    End If
    Set l = ParseBin(lvl + 1)
    Do
        op = BinOpAt(lvl)
        If op = 0 Then Exit Do
        Set nd = NewNode(NK_BINOP, tkLine(P))
        P = P + 1
        nd.Op = op
        Set nd.A = l
        Set nd.B = ParseBin(lvl + 1)
        Set l = nd
    Loop
    Set ParseBin = l
End Function

Private Function ParseFactor() As PyNode
    Dim nd As PyNode, ln As Long, op As Long, inner As PyNode, cv As Variant
    If tkType(P) = TT_OP Then
        Select Case tkText(P)
        Case "-": op = OP_NEG
        Case "+": op = OP_POS
        Case "~": op = OP_INVERT
        End Select
    End If
    If op = 0 Then
        Set ParseFactor = ParsePower()
        Exit Function
    End If
    ln = tkLine(P)
    P = P + 1
    Set inner = ParseFactor()
    If op = OP_NEG And inner.Kind = NK_CONST Then
        cv = inner.V
        If IsNum(cv) And VarType(cv) <> vbBoolean Then
            If IsFloatType(cv) Then
                inner.V = -CDbl(cv)
            Else
                inner.V = IntNeg(cv)
            End If
            Set ParseFactor = inner
            Exit Function
        End If
    End If
    Set nd = NewNode(NK_UNARY, ln)
    nd.Op = op
    Set nd.A = inner
    Set ParseFactor = nd
End Function

Private Function ParsePower() As PyNode
    Dim b As PyNode, nd As PyNode
    Set b = ParseAtomExpr()
    If IsOp("**") Then
        Set nd = NewNode(NK_BINOP, tkLine(P))
        P = P + 1
        nd.Op = OP_POW
        Set nd.A = b
        Set nd.B = ParseFactor()
        Set ParsePower = nd
    Else
        Set ParsePower = b
    End If
End Function

Private Function ParseAtomExpr() As PyNode
    Dim e As PyNode, nd As PyNode, ln As Long
    Set e = ParseAtom()
    Do
        If tkType(P) <> TT_OP Then Exit Do
        Select Case tkText(P)
        Case "("
            Set e = ParseCall(e)
        Case "["
            ln = tkLine(P)
            P = P + 1
            Set nd = NewNode(NK_SUBSCR, ln)
            Set nd.A = e
            Set nd.B = ParseSubscriptList()
            ExpectOp "]"
            Set e = nd
        Case "."
            ln = tkLine(P)
            P = P + 1
            Set nd = NewNode(NK_ATTR, ln)
            Set nd.A = e
            nd.S = ExpectName()
            nd.H = HashStr(nd.S)
            Set e = nd
        Case Else
            Exit Do
        End Select
    Loop
    Set ParseAtomExpr = e
End Function

Private Function ParseCall(ByVal fn As PyNode) As PyNode
    Dim nd As PyNode, a As PyNode, k As PyNode, ln As Long
    ln = tkLine(P)
    P = P + 1
    Set nd = NewNode(NK_CALL, ln)
    Set nd.A = fn
    Do While Not IsOp(")")
        If AcceptOp("**") Then
            Set k = NewNode(NK_DSTAR, ln)
            Set k.A = ParseTest()
            nd.Add k
        ElseIf AcceptOp("*") Then
            Set k = NewNode(NK_STARRED, ln)
            Set k.A = ParseTest()
            nd.Add k
        ElseIf tkType(P) = TT_NAME And PeekOpAt(1, "=") Then
            Set k = NewNode(NK_KWARG, ln)
            k.S = tkText(P)
            P = P + 2
            Set k.A = ParseTest()
            nd.Add k
        Else
            Set a = ParseTest()
            If IsKw("for") Then Set a = ParseComp(NK_GENEXP, a, Nothing, ln)
            nd.Add a
        End If
        If Not AcceptOp(",") Then Exit Do
    Loop
    ExpectOp ")"
    Set ParseCall = nd
End Function

Private Function ParseSubscriptList() As PyNode
    Dim first As PyNode, tup As PyNode
    Set first = ParseSubscript()
    If Not IsOp(",") Then
        Set ParseSubscriptList = first
        Exit Function
    End If
    Set tup = NewNode(NK_TUPLE, first.Ln)
    tup.Add first
    Do While AcceptOp(",")
        If IsOp("]") Then Exit Do
        tup.Add ParseSubscript()
    Loop
    Set ParseSubscriptList = tup
End Function

Private Function ParseSubscript() As PyNode
    Dim lo As PyNode, nd As PyNode, ln As Long
    ln = tkLine(P)
    If Not IsOp(":") Then
        Set lo = ParseTest()
        If Not IsOp(":") Then
            Set ParseSubscript = lo
            Exit Function
        End If
    End If
    Set nd = NewNode(NK_SLICE, ln)
    Set nd.A = lo
    P = P + 1
    If Not (IsOp("]") Or IsOp(",") Or IsOp(":")) Then Set nd.B = ParseTest()
    If AcceptOp(":") Then
        If Not (IsOp("]") Or IsOp(",")) Then Set nd.C = ParseTest()
    End If
    Set ParseSubscript = nd
End Function

Private Function ParseComp(ByVal kind As Long, ByVal elt As PyNode, ByVal valNode As PyNode, ByVal ln As Long) As PyNode
    Dim nd As PyNode, f As PyNode, c As PyNode
    Set nd = NewNode(kind, ln)
    Set nd.A = elt
    Set nd.B = valNode
    Do While IsKw("for")
        Set f = NewNode(NK_COMPFOR, tkLine(P))
        P = P + 1
        Set f.A = ParseTargetList()
        CheckTarget f.A, True
        ExpectKw "in"
        Set f.B = ParseOrTest()
        nd.Add f
        Do While IsKw("if")
            Set c = NewNode(NK_COMPIF, tkLine(P))
            P = P + 1
            Set c.A = ParseOrTest()
            nd.Add c
        Loop
    Loop
    Set ParseComp = nd
End Function

Private Function ParseAtom() As PyNode
    Dim nd As PyNode, ln As Long, t As String, first As PyNode
    ln = tkLine(P)
    Select Case tkType(P)
    Case TT_NUM
        Set nd = NewNode(NK_CONST, ln)
        nd.V = tkVal(P)
        P = P + 1
    Case TT_STR, TT_FSTR
        Set nd = ParseStrings()
    Case TT_NAME
        t = tkText(P)
        Select Case t
        Case "None"
            Set nd = NewNode(NK_CONST, ln)
            P = P + 1
        Case "True"
            Set nd = NewNode(NK_CONST, ln)
            nd.V = True
            P = P + 1
        Case "False"
            Set nd = NewNode(NK_CONST, ln)
            nd.V = False
            P = P + 1
        Case "yield"
            SynErr "yield expressions are only supported as statements"
        Case "await", "async"
            SynErr "async code is not supported"
        Case Else
            If IsReserved(t) Then SynErr "invalid syntax"
            Set nd = NameNode(t, ln)
            If t = "super" Then
                If Not (mCurDef Is Nothing) Then mCurDef.Op = mCurDef.Op Or 1
            End If
            P = P + 1
        End Select
    Case TT_OP
        t = tkText(P)
        Select Case t
        Case "("
            P = P + 1
            If AcceptOp(")") Then
                Set nd = NewNode(NK_TUPLE, ln)
            Else
                Set first = ParseTestOrStar()
                If IsKw("for") Then
                    Set nd = ParseComp(NK_GENEXP, first, Nothing, ln)
                ElseIf IsOp(",") Then
                    Set nd = NewNode(NK_TUPLE, ln)
                    nd.Add first
                    Do While AcceptOp(",")
                        If IsOp(")") Then Exit Do
                        nd.Add ParseTestOrStar()
                    Loop
                Else
                    Set nd = first
                End If
                ExpectOp ")"
            End If
        Case "["
            P = P + 1
            Set nd = NewNode(NK_LIST, ln)
            If Not IsOp("]") Then
                Set first = ParseTestOrStar()
                If IsKw("for") Then
                    Set nd = ParseComp(NK_LISTCOMP, first, Nothing, ln)
                Else
                    nd.Add first
                    Do While AcceptOp(",")
                        If IsOp("]") Then Exit Do
                        nd.Add ParseTestOrStar()
                    Loop
                End If
            End If
            ExpectOp "]"
        Case "{"
            P = P + 1
            Set nd = ParseBraces(ln)
            ExpectOp "}"
        Case "..."
            Set nd = NewNode(NK_CONST, ln)
            P = P + 1
        Case Else
            SynErr "invalid syntax"
        End Select
    Case TT_INDENT
        LexErr "unexpected indent", ln, True
    Case Else
        SynErr "invalid syntax"
    End Select
    Set ParseAtom = nd
End Function

Private Function ParseBraces(ByVal ln As Long) As PyNode
    Dim nd As PyNode, k As PyNode, v As PyNode, ds As PyNode
    Set nd = NewNode(NK_DICT, ln)
    If IsOp("}") Then
        Set ParseBraces = nd
        Exit Function
    End If
    If AcceptOp("**") Then
        Set ds = NewNode(NK_DSTAR, ln)
        Set ds.A = ParseBin(1)
        nd.Add ds
        nd.Add Nothing
    Else
        Set k = ParseTestOrStar()
        If Not IsOp(":") Then
            ' a set
            If IsKw("for") Then
                Set ParseBraces = ParseComp(NK_SETCOMP, k, Nothing, ln)
                Exit Function
            End If
            Set nd = NewNode(NK_SET, ln)
            nd.Add k
            Do While AcceptOp(",")
                If IsOp("}") Then Exit Do
                nd.Add ParseTestOrStar()
            Loop
            Set ParseBraces = nd
            Exit Function
        End If
        P = P + 1
        Set v = ParseTest()
        If IsKw("for") Then
            Set ParseBraces = ParseComp(NK_DICTCOMP, k, v, ln)
            Exit Function
        End If
        nd.Add k
        nd.Add v
    End If
    Do While AcceptOp(",")
        If IsOp("}") Then Exit Do
        If AcceptOp("**") Then
            Set ds = NewNode(NK_DSTAR, ln)
            Set ds.A = ParseBin(1)
            nd.Add ds
            nd.Add Nothing
        Else
            Set k = ParseTest()
            ExpectOp ":"
            nd.Add k
            nd.Add ParseTest()
        End If
    Loop
    Set ParseBraces = nd
End Function

Private Sub AddConstPart(ByVal nd As PyNode, ByVal s As String, ByVal ln As Long)
    Dim c As PyNode
    If Len(s) = 0 Then Exit Sub
    If nd.N > 0 Then
        If nd.Kid(nd.N - 1).Kind = NK_CONST Then
            nd.Kid(nd.N - 1).V = nd.Kid(nd.N - 1).V & s
            Exit Sub
        End If
    End If
    Set c = NewNode(NK_CONST, ln)
    c.V = s
    nd.Add c
End Sub

Private Function ParseStrings() As PyNode
    Dim nd As PyNode, ln As Long, anyF As Boolean, j As Long, s As String, t As String, rawF As Boolean
    ln = tkLine(P)
    j = P
    Do While tkType(j) = TT_STR Or tkType(j) = TT_FSTR
        If tkType(j) = TT_FSTR Then anyF = True
        j = j + 1
    Loop
    If Not anyF Then
        Do While tkType(P) = TT_STR
            s = s & tkText(P)
            P = P + 1
        Loop
        Set nd = NewNode(NK_CONST, ln)
        nd.V = s
    Else
        Set nd = NewNode(NK_FSTRING, ln)
        Do While tkType(P) = TT_STR Or tkType(P) = TT_FSTR
            If tkType(P) = TT_STR Then
                AddConstPart nd, tkText(P), ln
                P = P + 1
            Else
                t = tkText(P)
                rawF = (tkFlag(P) = 1)
                P = P + 1
                ParseFStringInto nd, t, rawF, ln
            End If
        Loop
    End If
    Set ParseStrings = nd
End Function

Private Sub FlushLit(ByVal nd As PyNode, ByRef lit As String, ByVal isRaw As Boolean, ByVal ln As Long)
    If Len(lit) = 0 Then Exit Sub
    If isRaw Then
        AddConstPart nd, lit, ln
    Else
        AddConstPart nd, Unescape(lit, ln), ln
    End If
    lit = vbNullString
End Sub

Private Sub ParseFStringInto(ByVal nd As PyNode, ByVal s As String, ByVal isRaw As Boolean, ByVal ln As Long)
    Dim i As Long, n As Long, c As String, lit As String, depth As Long, j As Long, k As Long
    Dim exprText As String, conv As String, spec As String, q As String, fv As PyNode, tr As String
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid$(s, i, 1)
        If c = "{" Then
            If Mid$(s, i + 1, 1) = "{" Then
                lit = lit & "{"
                i = i + 2
            Else
                FlushLit nd, lit, isRaw, ln
                j = i + 1
                depth = 0
                q = vbNullString
                Do While j <= n
                    c = Mid$(s, j, 1)
                    If Len(q) > 0 Then
                        If c = q Then q = vbNullString
                    ElseIf c = "'" Or c = """" Then
                        q = c
                    ElseIf c = "(" Or c = "[" Or c = "{" Then
                        depth = depth + 1
                    ElseIf c = ")" Or c = "]" Then
                        depth = depth - 1
                    ElseIf c = "}" Then
                        If depth = 0 Then Exit Do
                        depth = depth - 1
                    ElseIf depth = 0 And c = "!" And Mid$(s, j + 1, 1) <> "=" Then
                        Exit Do
                    ElseIf depth = 0 And c = ":" Then
                        Exit Do
                    End If
                    j = j + 1
                Loop
                If j > n Then LexErr "f-string: expecting '}'", ln, False
                exprText = Mid$(s, i + 1, j - i - 1)
                conv = vbNullString
                If Mid$(s, j, 1) = "!" Then
                    conv = Mid$(s, j + 1, 1)
                    j = j + 2
                End If
                spec = vbNullString
                If Mid$(s, j, 1) = ":" Then
                    k = j + 1
                    depth = 0
                    Do While k <= n
                        c = Mid$(s, k, 1)
                        If c = "{" Then
                            depth = depth + 1
                        ElseIf c = "}" Then
                            If depth = 0 Then Exit Do
                            depth = depth - 1
                        End If
                        k = k + 1
                    Loop
                    spec = Mid$(s, j + 1, k - j - 1)
                    j = k
                End If
                If Mid$(s, j, 1) <> "}" Then LexErr "f-string: expecting '}'", ln, False
                tr = RTrim$(exprText)
                If Right$(tr, 1) = "=" And Right$(tr, 2) <> "==" And Right$(tr, 2) <> "!=" And Right$(tr, 2) <> "<=" And Right$(tr, 2) <> ">=" Then
                    AddConstPart nd, exprText, ln
                    exprText = Left$(tr, Len(tr) - 1)
                    If Len(conv) = 0 And Len(spec) = 0 Then conv = "r"
                End If
                Set fv = NewNode(NK_FVALUE, ln)
                Set fv.A = ParseSubExpr(exprText, ln)
                fv.S = conv
                If Len(spec) > 0 Then
                    If InStr(1, spec, "{") > 0 Then
                        Set fv.B = NewNode(NK_FSTRING, ln)
                        ParseFStringInto fv.B, spec, True, ln
                    Else
                        fv.V = spec
                    End If
                End If
                nd.Add fv
                i = j + 1
            End If
        ElseIf c = "}" Then
            If Mid$(s, i + 1, 1) = "}" Then
                lit = lit & "}"
                i = i + 2
            Else
                LexErr "f-string: single '}' is not allowed", ln, False
            End If
        Else
            lit = lit & c
            i = i + 1
        End If
    Loop
    FlushLit nd, lit, isRaw, ln
End Sub

Private Function ParseSubExpr(ByVal text As String, ByVal ln As Long) As PyNode
    Dim saveP As Long, startTok As Long, e As PyNode
    If Len(Trim$(text)) = 0 Then LexErr "f-string: empty expression not allowed", ln, False
    saveP = P
    startTok = tkN
    Tokenize text, ln, True
    P = startTok
    Set e = ParseTestListStar()
    If tkType(P) <> TT_EOF Then LexErr "f-string: invalid syntax", ln, False
    P = saveP
    Set ParseSubExpr = e
End Function
