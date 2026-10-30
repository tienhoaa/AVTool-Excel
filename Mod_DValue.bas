Option Explicit

'=========================================================
' DVALUE - PHIEN BAN NANG CAP CHUYEN NGHIEP (FIX ALL CASES)
'
' Ho tro:
'   *  x  X  x
'   /  :
'   +  -  ^
'   ()  []
'   %
'   Dau thap phan , hoac .
'   Ngoac phep tinh long nhau & don vi o cuoi ngoac:
'     1 ben x 2 canh x (2 mat x 1,57m x 0,65m+0.2m x 0,4m)
'=========================================================

Public Function DValue(ByVal Expr As Variant) As Variant
    Dim s As String
    Dim Calc As String
    Dim i As Long
    Dim c As String
    Dim nxt As String
    Dim v As Variant
    Dim InUnit As Boolean
    Dim BracketLevel As Long

    On Error GoTo Loi
    Application.EnableCancelKey = xlDisabled

    If IsError(Expr) Then
        DValue = Expr
        Exit Function
    End If

    s = Trim$(CStr(Expr))
    If Len(s) = 0 Then
        DValue = 0
        Exit Function
    End If

    s = Replace(s, ChrW(&HD7), "*")
    s = Replace(s, ChrW(&HF7), "/")
    s = Replace(s, ChrW(&H2013), "-")
    s = Replace(s, ChrW(&H2014), "-")
    s = Replace(s, "[", "(")
    s = Replace(s, "]", ")")
    s = Replace(s, "{", "(")
    s = Replace(s, "}", ")")

    Calc = ""
    InUnit = False
    BracketLevel = 0

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)

        If InUnit Then
            Select Case c
                Case "*"
                    Calc = Calc & "*"
                    InUnit = False
                Case "x", "X"
                    Calc = Calc & "*"
                    InUnit = False
                Case "+"
                    Calc = Calc & "+"
                    InUnit = False
                Case "-"
                    Calc = Calc & "-"
                    InUnit = False
                Case "/"
                    If DValue_NextIsNumberOrBracket(s, i + 1) Then
                        Calc = Calc & "/"
                        InUnit = False
                    End If
                Case ":"
                    If DValue_NextIsNumberOrBracket(s, i + 1) Then
                        Calc = Calc & "/"
                        InUnit = False
                    End If
                Case "%"
                    Calc = Calc & "/100"
                    InUnit = False
                Case ")"
                    InUnit = False
                    If BracketLevel > 0 Then
                        Calc = Calc & ")"
                        BracketLevel = BracketLevel - 1
                    End If
                Case Else
                    ' Bo qua chu va so don vi
            End Select
        Else
            Select Case c
                Case "0" To "9"
                    Calc = Calc & c
                Case ",", "."
                    Calc = Calc & "."
                Case "+", "-"
                    Calc = Calc & c
                Case "*", "x", "X"
                    Calc = Calc & "*"
                Case "/", ":"
                    Calc = Calc & "/"
                Case "^"
                    Calc = Calc & "^"
                Case "%"
                    Calc = Calc & "/100"
                Case "("
                    If DValue_IsUnitBracket(s, i) Then
                        InUnit = True
                    Else
                        Calc = Calc & "("
                        BracketLevel = BracketLevel + 1
                    End If
                Case ")"
                    If BracketLevel > 0 Then
                        Calc = Calc & ")"
                        BracketLevel = BracketLevel - 1
                    End If
                Case " ", vbTab, "="
                    ' Bo qua
                Case Else
                    If DValue_IsUnitChar(c) Then
                        InUnit = True
                    End If
            End Select
        End If
    Next i

    Calc = Trim$(Calc)
    Do While Len(Calc) > 0
        c = Right$(Calc, 1)
        Select Case c
            Case "+", "-", "*", "/", "^", "."
                Calc = Left$(Calc, Len(Calc) - 1)
            Case Else
                Exit Do
        End Select
    Loop

    ' Tu dong dong cac ngoac phep tinh con thieu
    Do While BracketLevel > 0
        Calc = Calc & ")"
        BracketLevel = BracketLevel - 1
    Loop

    If Len(Calc) = 0 Then
        DValue = 0
        Exit Function
    End If

    v = Application.Evaluate(Calc)
    If IsError(v) Then
        DValue = CVErr(xlErrValue)
    Else
        DValue = v
    End If
    Exit Function

Loi:
    DValue = CVErr(xlErrValue)
End Function

Private Function DValue_NextIsNumberOrBracket(ByVal s As String, ByVal StartPos As Long) As Boolean
    Dim i As Long
    Dim c As String

    For i = StartPos To Len(s)
        c = Mid$(s, i, 1)
        If c <> " " And c <> vbTab Then
            Select Case c
                Case "0" To "9", "("
                    DValue_NextIsNumberOrBracket = True
                Case Else
                    DValue_NextIsNumberOrBracket = False
            End Select
            Exit Function
        End If
    Next i
    DValue_NextIsNumberOrBracket = False
End Function

Private Function DValue_IsUnitChar(ByVal c As String) As Boolean
    If Len(c) = 0 Then
        DValue_IsUnitChar = False
        Exit Function
    End If
    Select Case c
        Case "0" To "9"
            DValue_IsUnitChar = False
        Case "+", "-", "*", "/", "\", "^", "(", ")", "[", "]", "{", "}", ".", ",", ":", "%", "=", " ", vbTab, "x", "X"
            DValue_IsUnitChar = False
        Case Else
            DValue_IsUnitChar = True
    End Select
End Function

Private Function DValue_IsUnitBracket(ByVal s As String, ByVal StartPos As Long) As Boolean
    Dim i As Long
    Dim c As String
    Dim Level As Long
    Dim hasUnit As Boolean

    Level = 1
    hasUnit = False

    For i = StartPos + 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = "(" Then
            Level = Level + 1
        ElseIf c = ")" Then
            Level = Level - 1
            If Level = 0 Then Exit For
        Else
            ' Neu chua toan tu toan hoc -> CHAC CHAN LA NGOAC PHEP TINH
            Select Case c
                Case "+", "-", "*", "^", "x", "X"
                    DValue_IsUnitBracket = False
                    Exit Function
                Case "/"
                    If DValue_NextIsNumberOrBracket(s, i + 1) Then
                        DValue_IsUnitBracket = False
                        Exit Function
                    End If
            End Select

            If DValue_IsUnitChar(c) Then
                hasUnit = True
            End If
        End If
    Next i

    DValue_IsUnitBracket = hasUnit
End Function
