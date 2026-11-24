VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PyNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' PyNode - one node of the abstract syntax tree produced by PyParser.
' The meaning of the generic fields depends on Kind (see NK_* in PyRuntime).
' ---------------------------------------------------------------------------
Public Kind As Long         ' NK_* node kind
Public Op As Long           ' operator code / flags
Public S As String          ' name, attribute name, string data
Public V As Variant         ' constant value / extra data
Public H As Long            ' cached hash of S (names)
Public A As PyNode          ' generic children
Public B As PyNode
Public C As PyNode
Public D As PyNode
Public Ln As Long           ' source line
Public N As Long            ' number of list children
Private mKids() As PyNode

Public Sub Add(ByVal k As PyNode)
    If N = 0 Then
        ReDim mKids(0 To 3)
    ElseIf N > UBound(mKids) Then
        ReDim Preserve mKids(0 To 2 * N - 1)
    End If
    Set mKids(N) = k
    N = N + 1
End Sub

Public Function Kid(ByVal i As Long) As PyNode
    Set Kid = mKids(i)
End Function

Public Sub SetKid(ByVal i As Long, ByVal k As PyNode)
    Set mKids(i) = k
End Sub

Public Sub RemoveLastKid()
    If N > 0 Then
        N = N - 1
        Set mKids(N) = Nothing
    End If
End Sub
