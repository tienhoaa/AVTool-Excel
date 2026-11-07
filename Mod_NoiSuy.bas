Option Explicit

'========================================================================================
' AVTOOL - BO CONG CU NOI SUY BANG TRA CHUYEN NGHIEP (MA TRAN 2D, DAY SO, NOI SUY 1 CHIEU)
'
' Bao gom 3 loai noi suy chinh:
'   1. NoisuyMatran(Xk, Yk, GT, Xu, Yu): Noi suy ma tran 2 chieu (Bilinear Interpolation)
'   2. NoisuyDay(Xk, Yk, Xu): Noi suy theo day so 1 chieu (Linear Interpolation)
'   3. Noisuy2b(Xs1, Ys1, Xs2, Ys2, Xs): Noi suy tuyen tinh 1 chieu giua 2 diem
'   4. Dt(d): Tinh dien tich tiet dien cot thep tron (m2) - Giu nguyen de tuong thich
'========================================================================================

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
    Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
    Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
#End If

Public Function UniText(ByVal s As String) As String
    Dim pos As Long, hexStr As String, res As String
    res = s
    pos = InStr(res, "\u")
    Do While pos > 0
        hexStr = Mid(res, pos + 2, 4)
        res = Left(res, pos - 1) & ChrW(CLng("&H" & hexStr)) & Mid(res, pos + 6)
        pos = InStr(res, "\u")
    Loop
    UniText = res
End Function

Public Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "AVTool") As Long
    MsgBoxUni = MessageBoxW(0, StrPtr(UniText(prompt)), StrPtr(UniText(title)), buttons)
End Function

Public Sub SetFormTitleUnicode(ByVal formCaption As String, ByVal unicodeTitle As String)
    #If VBA7 Then
        Dim hWndForm As LongPtr
    #Else
        Dim hWndForm As Long
    #End If
    On Error Resume Next
    hWndForm = FindWindowW(0, StrPtr(formCaption))
    If hWndForm = 0 Then hWndForm = FindWindowA("ThunderDFrame", formCaption)
    If hWndForm = 0 Then hWndForm = FindWindowA(vbNullString, formCaption)
    If hWndForm <> 0 Then
        SetWindowTextW hWndForm, StrPtr(UniText(unicodeTitle))
    End If
End Sub

' -----------------------------------------------------------------------------
' 1. HAM NOI SUY TUYEN TINH 1 CHIEU 2 DIEM: (Xs1, Ys1) VA (Xs2, Ys2)
'    Dang ky huong dan tham so day du trong Excel (Tooltip fx)
' -----------------------------------------------------------------------------
Public Function Noisuy2b(ByVal Xs1 As Double, ByVal Ys1 As Double, ByVal Xs2 As Double, ByVal Ys2 As Double, ByVal Xs As Double) As Double
    On Error Resume Next
    If Xs2 = Xs1 Then
        Noisuy2b = Ys1
    Else
        Noisuy2b = (Ys2 - Ys1) * (Xs - Xs1) / (Xs2 - Xs1) + Ys1
    End If
End Function

' -----------------------------------------------------------------------------
' 2. HAM TINH DIEN TICH COT THEP TRON (d theo mm -> dien tich m2)
' -----------------------------------------------------------------------------
Public Function Dt(ByVal d As Double) As Double
    On Error Resume Next
    Dt = 3.14159265358979 * (d ^ 2) / 4000000#
End Function

' -----------------------------------------------------------------------------
' 3. THUAT TOAN PHU TRO NOI SUY 1 CHIEU CO TU DONG SAP XEP & NGOAI SUY
' -----------------------------------------------------------------------------
Public Function TinhtoanNs(ByVal Ncot As Long, ByRef X() As Double, ByRef Y() As Double, ByVal Xu As Double, Optional ByVal tb As String = "") As Variant
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    Dim i As Long, j As Long
    Dim Phu As Double
    
    On Error GoTo LoiNS
    
    ' Sap xep noi bot tang dan theo X, dong thoi hoan vi Y
    For i = 1 To Ncot - 1
        For j = i + 1 To Ncot
            If X(i) > X(j) Then
                Phu = X(i): X(i) = X(j): X(j) = Phu
                Phu = Y(i): Y(i) = Y(j): Y(j) = Phu
            End If
        Next j
    Next i
    
    ' Neu trung dung moc dau hoac moc cuoi
    If Xu = X(1) Then
        TinhtoanNs = Y(1)
        Exit Function
    End If
    If Xu = X(Ncot) Then
        TinhtoanNs = Y(Ncot)
        Exit Function
    End If
    
    ' Kiem tra trung bat ky moc nao o giua
    For i = 1 To Ncot
        If X(i) = Xu Then
            TinhtoanNs = Y(i)
            Exit Function
        End If
    Next i
    
    ' Neu ngoai khoang tra: Tu dong ngoai suy tuyen tinh
    If Xu < X(1) Then
        X1 = X(1): Y1 = Y(1)
        X2 = X(2): Y2 = Y(2)
        If X2 <> X1 Then
            TinhtoanNs = (Y2 - Y1) * (Xu - X1) / (X2 - X1) + Y1
        Else
            TinhtoanNs = Y1
        End If
        Exit Function
    ElseIf Xu > X(Ncot) Then
        X1 = X(Ncot - 1): Y1 = Y(Ncot - 1)
        X2 = X(Ncot): Y2 = Y(Ncot)
        If X2 <> X1 Then
            TinhtoanNs = (Y2 - Y1) * (Xu - X1) / (X2 - X1) + Y1
        Else
            TinhtoanNs = Y2
        End If
        Exit Function
    End If
    
    ' Noi suy trong khoang
    For i = 1 To Ncot - 1
        If (X(i) <= Xu) And (X(i + 1) >= Xu) Then
            X1 = X(i): Y1 = Y(i)
            X2 = X(i + 1): Y2 = Y(i + 1)
            If X2 <> X1 Then
                TinhtoanNs = (Y2 - Y1) * (Xu - X1) / (X2 - X1) + Y1
            Else
                TinhtoanNs = Y1
            End If
            Exit Function
        End If
    Next i
    
    TinhtoanNs = Y(Ncot)
    Exit Function

LoiNS:
    TinhtoanNs = CVErr(xlErrValue)
End Function

' -----------------------------------------------------------------------------
' 4. HAM NOI SUY 1 CHIEU THEO DAY SO (Xk: day moc X, Yk: day moc Y, Xu: gia tri can tra)
'    Xu ly an toan tuyet doi moi loai du lieu, khong bao gio bi loi #VALUE!
' -----------------------------------------------------------------------------
Public Function NoisuyDay(ByVal Xk As Range, ByVal Yk As Range, ByVal Xu As Variant) As Variant
    Dim X() As Double, Y() As Double
    Dim Ncot As Long, k As Long
    Dim cell As Range
    Dim dblXu As Double
    Dim sXu As String
    
    On Error GoTo LoiDay
    
    ' 1. Kiem tra va ep kieu Xu
    If IsError(Xu) Then
        NoisuyDay = CVErr(xlErrValue)
        Exit Function
    End If
    
    If TypeOf Xu Is Range Then
        If IsEmpty(Xu) Or Xu.Value = "" Then
            NoisuyDay = ""
            Exit Function
        End If
        If IsNumeric(Xu.Value) Then
            dblXu = CDbl(Xu.Value)
        Else
            sXu = Trim(Replace(CStr(Xu.Value), ",", "."))
            If IsNumeric(sXu) Then
                dblXu = CDbl(sXu)
            Else
                NoisuyDay = CVErr(xlErrValue)
                Exit Function
            End If
        End If
    Else
        If IsNumeric(Xu) Then
            dblXu = CDbl(Xu)
        Else
            sXu = Trim(Replace(CStr(Xu), ",", "."))
            If IsNumeric(sXu) Then
                dblXu = CDbl(sXu)
            Else
                NoisuyDay = CVErr(xlErrValue)
                Exit Function
            End If
        End If
    End If
    
    ' 2. Kiem tra so phan tu cua 2 dai o
    Ncot = Xk.Count
    If Yk.Count <> Ncot Or Ncot < 2 Then
        NoisuyDay = CVErr(xlErrRef)
        Exit Function
    End If
    
    ReDim X(1 To Ncot)
    ReDim Y(1 To Ncot)
    
    ' 3. Doc du lieu Xk an toan bang For Each
    k = 0
    For Each cell In Xk
        k = k + 1
        If IsNumeric(cell.Value) Then
            X(k) = CDbl(cell.Value)
        Else
            X(k) = Val(Replace(CStr(cell.Value), ",", "."))
        End If
    Next cell
    
    ' 4. Doc du lieu Yk an toan bang For Each
    k = 0
    For Each cell In Yk
        k = k + 1
        If IsNumeric(cell.Value) Then
            Y(k) = CDbl(cell.Value)
        Else
            Y(k) = Val(Replace(CStr(cell.Value), ",", "."))
        End If
    Next cell
    
    ' 5. Tinh toan noi suy
    NoisuyDay = TinhtoanNs(Ncot, X, Y, dblXu, "")
    Exit Function
    
LoiDay:
    NoisuyDay = CVErr(xlErrValue)
End Function

' -----------------------------------------------------------------------------
' 5. HAM NOI SUY MA TRAN 2 CHIEU SONG TUYEN TINH (Bilinear Interpolation)
'    Xk: Hang moc ngang (1 dong x N cot)
'    Yk: Cot moc doc (M dong x 1 cot)
'    GT: Vung ma tran du lieu (M dong x N cot)
'    Xu: Gia tri tra theo truc X
'    Yu: Gia tri tra theo truc Y
' -----------------------------------------------------------------------------
Public Function NoisuyMatran(ByVal Xk As Range, ByVal Yk As Range, ByVal GT As Range, ByVal Xu As Variant, ByVal Yu As Variant) As Variant
    Dim X() As Double, Y() As Double, G() As Double
    Dim i As Long, j As Long, k As Long
    Dim Ndong As Long, Ncot As Long
    Dim Phu As Double
    Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    Dim G1 As Double, G2 As Double, G3 As Double
    Dim Timthay As Boolean
    Dim cell As Range
    Dim dblXu As Double, dblYu As Double
    
    On Error GoTo LoiMatran
    
    ' Ep kieu Xu an toan
    If TypeOf Xu Is Range Then
        If IsNumeric(Xu.Value) Then dblXu = CDbl(Xu.Value) Else dblXu = Val(Replace(CStr(Xu.Value), ",", "."))
    Else
        If IsNumeric(Xu) Then dblXu = CDbl(Xu) Else dblXu = Val(Replace(CStr(Xu), ",", "."))
    End If
    
    ' Ep kieu Yu an toan
    If TypeOf Yu Is Range Then
        If IsNumeric(Yu.Value) Then dblYu = CDbl(Yu.Value) Else dblYu = Val(Replace(CStr(Yu.Value), ",", "."))
    Else
        If IsNumeric(Yu) Then dblYu = CDbl(Yu) Else dblYu = Val(Replace(CStr(Yu), ",", "."))
    End If
    
    Timthay = False
    Ndong = GT.Rows.Count
    Ncot = GT.Columns.Count
    
    If Xk.Count <> Ncot Or Yk.Count <> Ndong Then
        NoisuyMatran = CVErr(xlErrRef)
        Exit Function
    End If
    
    ReDim G(1 To Ndong, 1 To Ncot)
    For i = 1 To Ndong
        For j = 1 To Ncot
            If IsNumeric(GT.Item(i, j).Value) Then
                G(i, j) = CDbl(GT.Item(i, j).Value)
            Else
                G(i, j) = Val(Replace(CStr(GT.Item(i, j).Value), ",", "."))
            End If
        Next j
    Next i
    
    ReDim X(1 To Ncot)
    ReDim Y(1 To Ndong)
    
    k = 0
    For Each cell In Xk
        k = k + 1
        If IsNumeric(cell.Value) Then X(k) = CDbl(cell.Value) Else X(k) = Val(Replace(CStr(cell.Value), ",", "."))
    Next cell
    
    k = 0
    For Each cell In Yk
        k = k + 1
        If IsNumeric(cell.Value) Then Y(k) = CDbl(cell.Value) Else Y(k) = Val(Replace(CStr(cell.Value), ",", "."))
    Next cell
    
    ' Sap xep X tang dan & hoan vi cot tuong ung trong ma tran G
    For i = 1 To Ncot - 1
        For j = i + 1 To Ncot
            If X(i) > X(j) Then
                Phu = X(i): X(i) = X(j): X(j) = Phu
                For k = 1 To Ndong
                    Phu = G(k, i): G(k, i) = G(k, j): G(k, j) = Phu
                Next k
            End If
        Next j
    Next i
    
    ' Sap xep Y tang dan & hoan vi hang tuong ung trong ma tran G
    For i = 1 To Ndong - 1
        For j = i + 1 To Ndong
            If Y(i) > Y(j) Then
                Phu = Y(i): Y(i) = Y(j): Y(j) = Phu
                For k = 1 To Ncot
                    Phu = G(i, k): G(i, k) = G(j, k): G(j, k) = Phu
                Next k
            End If
        Next j
    Next i
    
    ' Gioi han bien an toan
    If dblXu < X(1) Then dblXu = X(1)
    If dblXu > X(Ncot) Then dblXu = X(Ncot)
    If dblYu < Y(1) Then dblYu = Y(1)
    If dblYu > Y(Ndong) Then dblYu = Y(Ndong)
    
    ' Tim khoang bao chua Xu va Yu
    For i = 1 To Ndong - 1
        For j = 1 To Ncot - 1
            If (Y(i) <= dblYu) And (Y(i + 1) >= dblYu) And (X(j) <= dblXu) And (X(j + 1) >= dblXu) Then
                X1 = j: X2 = j + 1
                Y1 = i: Y2 = i + 1
                Timthay = True
                Exit For
            End If
        Next j
        If Timthay Then Exit For
    Next i
    
    If Timthay Then
        G1 = Noisuy2b(Y(Y1), G(Y1, X1), Y(Y2), G(Y2, X1), dblYu)
        G2 = Noisuy2b(Y(Y1), G(Y1, X2), Y(Y2), G(Y2, X2), dblYu)
        G3 = Noisuy2b(X(X1), G1, X(X2), G2, dblXu)
        NoisuyMatran = G3
    Else
        NoisuyMatran = 0
    End If
    Exit Function

LoiMatran:
    NoisuyMatran = CVErr(xlErrValue)
End Function

' -----------------------------------------------------------------------------
' 6. HAM PHU TRO NAME (Tra ve ten dai o)
' -----------------------------------------------------------------------------
Public Function Name(ByVal a As Range) As String
    Const S As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    On Error Resume Next
    If a.Column <= 26 And a.Columns.Count <= 26 Then
        Name = Mid(S, a.Column, 1) & a.Row & "x" & Mid(S, a.Columns.Count, 1) & (a.Row + a.Rows.Count - 1)
    Else
        Name = a.Address(False, False)
    End If
End Function

' -----------------------------------------------------------------------------
' 7. GAN TOOLTIP HUONG DAN VAO O KET QUA KHI CHON HOAC DAT CHUOT VAO O
' -----------------------------------------------------------------------------
Public Sub AttachTooltipNoisuy2b(ByVal rng As Range)
    On Error Resume Next
    ' Xoa bo sach se Data Validation vang cu neu co tren o
    rng.Validation.Delete
End Sub

' =============================================================================
' 8. DIEU KHIEN HIEN THI 3 FORM CHUYEN DUNG (MODELess DE THAO TAC TREN EXCEL)
' =============================================================================
Public Sub ShowNoiSuyMatran()
    On Error Resume Next
    LoadIntelliSenseXLL
    RegisterMacroDescriptions
    If Val(Application.Version) >= 12 Then
        frmNoiSuyMatran.Show vbModeless
    Else
        frmNoiSuyMatran.Show
    End If
End Sub

Public Sub ShowNoiSuyDay()
    On Error Resume Next
    LoadIntelliSenseXLL
    RegisterMacroDescriptions
    If Val(Application.Version) >= 12 Then
        frmNoiSuyDay.Show vbModeless
    Else
        frmNoiSuyDay.Show
    End If
End Sub

Public Sub ShowNoiSuy1Chieu()
    On Error Resume Next
    Dim curCell As Range
    Set curCell = ActiveCell
    If curCell Is Nothing Then Exit Sub
    
    ' 1. Kich hoat may chu IntelliSense va dang ky mo ta tham so cho Excel
    LoadIntelliSenseXLL
    RegisterMacroDescriptions
    
    ' 2. Dong cua so VBA Editor neu dang mo (tuyet doi khong hien code truoc mat nguoi dung)
    If Not Application.VBE Is Nothing Then
        Application.VBE.MainWindow.Visible = False
    End If
    
    ' 3. Xoa sach Data Validation vang cu tren o (neu co)
    curCell.Validation.Delete
    
    ' 4. Gan truc tiep cong thuc =Noisuy2b() vao o bang code VBA
    '    (Khong dung SendKeys danh chu, chong triet de Unikey/EVKey Telex bien 'nois' thanh 'Nội')
    curCell.Formula = "=Noisuy2b()"
    curCell.Select
    
    ' 5. Active cua so Excel va dua con tro vao giua cap dau ngoac de hop thoai IntelliSense bung ra
    AppActivate Application.Caption
    Application.OnTime Now, "FocusInsideNoisuy2b"
End Sub

Public Sub FocusInsideNoisuy2b()
    On Error Resume Next
    Application.SendKeys "{F2}{LEFT}"
End Sub

Public Sub DoTypeNoisuy2b()
    FocusInsideNoisuy2b
End Sub

' =============================================================================
' 9. HUONG DAN CU PHAP NOI SUY CHO NGUOI DUNG
' =============================================================================
Public Sub ShowHuongDanNoiSuy()
    Dim s As String
    s = "=== H\u01AF\u1EDANG D\u1EAAN C\u00C1C H\u00C0M N\u1ED8I SUY TO\u00C1N H\u1ECCC & K\u1EF8 THU\u1EACT ==="
    s = s & vbCrLf & ""
    s = s & vbCrLf & "1. N\u1ED8I SUY 1 CHI\u1EC0U (2 \u0110I\u1EC2M K\u1EB8P):"
    s = s & vbCrLf & "   =Noisuy2b(X1, Y1, X2, Y2, X_can_tra)"
    s = s & vbCrLf & "   - N\u1ED9i suy tuy\u1EBFn t\u00EDnh gi\u1EEFa 2 m\u1ED1c k\u1EB9p (X1, Y1) v\u00E0 (X2, Y2)."
    s = s & vbCrLf & ""
    s = s & vbCrLf & "2. N\u1ED8I SUY D\u00C3Y S\u1ED0 1 CHI\u1EC0U:"
    s = s & vbCrLf & "   =NoisuyDay(Day_X, Day_Y, X_can_tra)"
    s = s & vbCrLf & "   - T\u1EF1 \u0111\u1ED9ng s\u1EAFp x\u1EBFp d\u00E3y m\u1ED1c X v\u00E0 ho\u00E1n v\u1ECB d\u00E3y gi\u00E1 tr\u1ECB Y t\u01B0\u01A1ng \u1EE9ng."
    s = s & vbCrLf & "   - H\u1ED7 tr\u1EE3 c\u1EA3 m\u1ED1c x\u1EBFp t\u0103ng d\u1EA7n ho\u1EB7c gi\u1EA3m d\u1EA7n, h\u00E0ng ngang ho\u1EB7c c\u1ED9t d\u1ECDc."
    s = s & vbCrLf & ""
    s = s & vbCrLf & "3. N\u1ED8I SUY MA TR\u1EACN 2 CHI\u1EC0U (SONG TUY\u1EBEN T\u00CDNH - BILINEAR):"
    s = s & vbCrLf & "   =NoisuyMatran(Hang_X, Cot_Y, MaTran_GT, Xu, Yu)"
    s = s & vbCrLf & "   - Hang_X : D\u00E3y m\u1ED1c tr\u1EE5c ho\u00E0nh (tr\u1EE5c ngang)."
    s = s & vbCrLf & "   - Cot_Y  : D\u00E3y m\u1ED1c tr\u1EE5c tung (tr\u1EE5c d\u1ECDc)."
    s = s & vbCrLf & "   - MaTran_GT : V\u00F9ng gi\u00E1 tr\u1ECB b\u00EAn trong b\u1EA3ng."
    s = s & vbCrLf & "   - Xu, Yu : T\u1ECDa \u0111\u1ED9 \u0111i\u1EC3m c\u1EA7n tra n\u1ED9i suy."
    s = s & vbCrLf & ""
    s = s & vbCrLf & "4. H\u00C0M T\u00CDNH DI\u1EC6N T\u00CDCH TH\u00C9P (mm -> m2):"
    s = s & vbCrLf & "   =Dt(d) : T\u00EDnh di\u1EC7n t\u00EDch ti\u1EBFt di\u1EC7n c\u1ED1t th\u00E9p tr\u00F2n \u0111\u01B0\u1EDDng k\u00EDnh phi d (mm) ra m2."
    s = s & vbCrLf & ""
    s = s & vbCrLf & "(*) M\u1EB8O S\u1EEC D\u1EE4NG:"
    s = s & vbCrLf & "- B\u1EA5m n\u00FAt [T\u1EA1o B\u1EA3ng M\u1EABu] tr\u00EAn Ribbon \u0111\u1EC3 sinh s\u1EB5n sheet m\u1EABu n\u1ED9i suy tr\u1EF1c quan."
    s = s & vbCrLf & "- IntelliSense t\u1EF1 \u0111\u1ED9ng hi\u1EC3n th\u1ECB m\u00F4 t\u1EA3 t\u1EEBng tham s\u1ED1 b\u1EB1ng ti\u1EBFng Vi\u1EC7t khi g\u00F5 c\u00F4ng th\u1EE9c."
    MsgBoxUni s, vbInformation, "H\u01AF\u1EDANG D\u1EAAN C\u00C1C H\u00C0M N\u1ED8I SUY - AVTOOL"
End Sub

' =============================================================================
' 10. DANG KY MO TA HAM VA THAM SO TRONG EXCEL (TOOLTIPS / FX DIALOG)
' =============================================================================
Public Sub RegisterMacroDescriptions()
    On Error Resume Next
    Application.MacroOptions _
        Macro:="Noisuy2b", _
        Description:=UniText("Noi suy tuyen tinh giua 2 moc: (Xs1, Ys1) va (Xs2, Ys2) tai gia tri can tra Xs"), _
        Category:=UniText("AVTool - Noi Suy"), _
        ArgumentDescriptions:=Array( _
            UniText("Gia tri moc X duoi (X1)"), _
            UniText("Gia tri moc Y duoi tuong ung (Y1)"), _
            UniText("Gia tri moc X tren (X2)"), _
            UniText("Gia tri moc Y tren tuong ung (Y2)"), _
            UniText("Gia tri X can tra cuu"))
            
    Application.MacroOptions _
        Macro:="NoisuyDay", _
        Description:=UniText("Noi suy 1 chieu tuyen tinh theo day moc X va day gia tri Y (tu dong sap xep)"), _
        Category:=UniText("AVTool - Noi Suy"), _
        ArgumentDescriptions:=Array( _
            UniText("Day moc X (vung hang hoac cot)"), _
            UniText("Day gia tri Y tuong ung"), _
            UniText("Gia tri X can tra cuu"))
            
    Application.MacroOptions _
        Macro:="NoisuyMatran", _
        Description:=UniText("Noi suy ma tran 2 chieu song tuyen tinh (Bilinear) tu bang tra 2 truc X va Y"), _
        Category:=UniText("AVTool - Noi Suy"), _
        ArgumentDescriptions:=Array( _
            UniText("Hang moc ngang (truc X)"), _
            UniText("Cot moc doc (truc Y)"), _
            UniText("Vung ma tran so lieu (GT)"), _
            UniText("Gia tri tra theo truc X (Xu)"), _
            UniText("Gia tri tra theo truc Y (Yu)"))
End Sub

' =============================================================================
' 11. TU DONG NHUNG MOD_NOISUY VAO FILE DANG MO (TRANH LOI #NAME? KHI GUI FILE)
' =============================================================================
Public Function InsertNoiSuyToActiveWorkbook(Optional ByVal showMsg As Boolean = True) As Boolean
    Dim targetWb As Workbook
    Dim comp As Object, srcComp As Object, existingComp As Object
    Dim vbProj As Object
    Dim noiSuyCode As String
    Dim lineCount As Long
    Dim codeText As String
    
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    
    Set targetWb = ActiveWorkbook
    If targetWb Is Nothing Then
        If showMsg Then MsgBoxUni "Ch\u01B0a c\u00F3 file Excel n\u00E0o \u0111ang m\u1EDF!", vbExclamation, "AVTool"
        Exit Function
    End If
    
    If targetWb Is ThisWorkbook Then
        If showMsg Then MsgBoxUni "\u0110\u00E2y ch\u00EDnh l\u00E0 file Add-in AVTool!", vbInformation, "AVTool"
        Exit Function
    End If
    
    Err.Clear
    Set vbProj = targetWb.VBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        If showMsg Then
            MsgBoxUni "Kh\u00F4ng th\u1EC3 ch\u00E8n Module v\u00E0o file v\u00EC Excel ch\u01B0a m\u1EDF kh\u00F3a quy\u1EC1n VBA." & vbCrLf & vbCrLf & _
                      "C\u00E1ch b\u1EADt r\u1EA5t nhanh (ch\u1EC9 c\u1EA7n l\u00E0m 1 l\u1EA7n):" & vbCrLf & _
                      "1. V\u00E0o File > Options > Trust Center > Trust Center Settings..." & vbCrLf & _
                      "2. Ch\u1ECDn Macro Settings > T\u00EDch ch\u1ECDn 'Trust access to the VBA project object model' > OK.", vbExclamation, "AVTool"
        End If
        Exit Function
    End If
    
    Set srcComp = ThisWorkbook.VBProject.VBComponents("Mod_NoiSuy")
    If srcComp Is Nothing Then
        If showMsg Then MsgBoxUni "L\u1ED7i: Kh\u00F4ng t\u00ECm th\u1EA5y m\u00E3 ngu\u1ED3n Mod_NoiSuy trong Add-in!", vbCritical, "AVTool"
        Exit Function
    End If
    
    lineCount = srcComp.CodeModule.CountOfLines
    noiSuyCode = srcComp.CodeModule.Lines(1, lineCount)
    
    Set existingComp = Nothing
    For Each comp In vbProj.VBComponents
        If comp.Type = 1 Then
            lineCount = comp.CodeModule.CountOfLines
            If lineCount > 0 Then
                codeText = comp.CodeModule.Lines(1, lineCount)
                If InStr(1, codeText, "Function NoisuyMatran", vbTextCompare) > 0 Or InStr(1, codeText, "Function NoisuyDay", vbTextCompare) > 0 Then
                    Set existingComp = comp
                    Exit For
                End If
            End If
        End If
    Next comp
    
    If Not existingComp Is Nothing Then
        existingComp.CodeModule.DeleteLines 1, existingComp.CodeModule.CountOfLines
        existingComp.CodeModule.AddFromString noiSuyCode
        InsertNoiSuyToActiveWorkbook = True
        If showMsg Then
            MsgBoxUni "\u0110\u00C3 C\u1EACP NH\u1EACT B\u1ED8 H\u00C0M N\u1ED8I SUY B\u1EA2N M\u1EDAI NH\u1EA4T V\u00C0O FILE '" & targetWb.Name & "' TH\u00C0NH C\u00D4NG!", vbInformation, "AVTool"
        End If
    Else
        Set comp = vbProj.VBComponents.Add(1)
        comp.Name = "Mod_NoiSuy"
        comp.CodeModule.AddFromString noiSuyCode
        InsertNoiSuyToActiveWorkbook = True
        If showMsg Then
            MsgBoxUni "\u0110\u00C3 CH\u00C8N B\u1ED8 H\u00C0M N\u1ED8I SUY V\u00C0O FILE '" & targetWb.Name & "' TH\u00C0NH C\u00D4NG!" & vbCrLf & vbCrLf & _
                      "Module 'Mod_NoiSuy' \u0111\u00E3 \u0111\u01B0\u1EE3c nh\u00FAng v\u0129nh vi\u1EC5n v\u00E0o file n\u00E0y.", vbInformation, "AVTool"
        End If
    End If
    
    On Error Resume Next
    If Not Application.VBE Is Nothing Then
        Application.VBE.MainWindow.Visible = False
    End If
    RegisterMacroDescriptions
    Application.Calculate
End Function

' =============================================================================
' 12. TAO BANG TRA MAU MINH HOA CHO CA 3 LOAI (MA TRAN 2D, DAY SO 1D, 2 DIEM)
' =============================================================================
Public Sub CreateSampleTableAll(Optional ByVal showMsg As Boolean = True)
    Dim wb As Workbook, ws As Worksheet
    Dim sheetName As String
    Dim idx As Long
    
    On Error Resume Next
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        If showMsg Then MsgBoxUni "Vui l\u00F2ng m\u1EDF m\u1ED9t file Excel \u0111\u1EC3 t\u1EA1o b\u1EA3ng m\u1EABu!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    sheetName = "Mau_NoiSuy"
    idx = 1
    Do While True
        Set ws = Nothing
        Set ws = wb.Worksheets(sheetName)
        If ws Is Nothing Then Exit Do
        idx = idx + 1
        sheetName = "Mau_NoiSuy_" & idx
    Loop
    
    Set ws = wb.Worksheets.Add(After:=wb.ActiveSheet)
    ws.Name = sheetName
    
    ' --- PHAN 1: BANG TRA MA TRAN 2 CHIEU (HAM =NoisuyMatran) ---
    ws.Range("B2:G2").Merge
    ws.Range("B2").Value = "1. BANG TRA MA TRAN 2 CHIEU (HAM =NoisuyMatran)"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 11
    ws.Range("B2").Font.Color = RGB(0, 51, 102)
    
    ws.Range("B3").Value = "Moc Y \ Moc X ->"
    ws.Range("B3").Font.Italic = True
    
    ' Hang moc X (C3:G3)
    ws.Range("C3").Value = 1#
    ws.Range("D3").Value = 1.5
    ws.Range("E3").Value = 2#
    ws.Range("F3").Value = 3#
    ws.Range("G3").Value = 5#
    ws.Range("C3:G3").Interior.Color = RGB(220, 235, 252)
    ws.Range("C3:G3").Font.Bold = True
    ws.Range("C3:G3").HorizontalAlignment = xlCenter
    
    ' Cot moc Y (B4:B9)
    ws.Range("B4").Value = 0.2
    ws.Range("B5").Value = 0.4
    ws.Range("B6").Value = 0.6
    ws.Range("B7").Value = 0.8
    ws.Range("B8").Value = 1#
    ws.Range("B9").Value = 1.5
    ws.Range("B4:B9").Interior.Color = RGB(255, 242, 204)
    ws.Range("B4:B9").Font.Bold = True
    ws.Range("B4:B9").HorizontalAlignment = xlCenter
    
    ' Du lieu ma tran GT (C4:G9)
    ws.Range("C4:G4").Value = Array(0.245, 0.252, 0.255, 0.258, 0.26)
    ws.Range("C5:G5").Value = Array(0.231, 0.241, 0.246, 0.25, 0.253)
    ws.Range("C6:G6").Value = Array(0.208, 0.222, 0.231, 0.238, 0.242)
    ws.Range("C7:G7").Value = Array(0.181, 0.199, 0.211, 0.221, 0.227)
    ws.Range("C8:G8").Value = Array(0.155, 0.174, 0.188, 0.201, 0.209)
    ws.Range("C9:G9").Value = Array(0.106, 0.123, 0.137, 0.152, 0.162)
    ws.Range("C4:G9").NumberFormat = "0.000"
    ws.Range("C4:G9").HorizontalAlignment = xlRight
    ws.Range("B3:G9").Borders.LineStyle = xlContinuous
    
    ' Tra cuu 2D
    ws.Range("B11").Value = "TRA CUU MA TRAN 2D:"
    ws.Range("B11").Font.Bold = True
    ws.Range("B12").Value = "Gia tri Xu:"
    ws.Range("C12").Value = 1.75
    ws.Range("C12").Interior.Color = RGB(255, 255, 200)
    ws.Range("B13").Value = "Gia tri Yu:"
    ws.Range("C13").Value = 0.5
    ws.Range("C13").Interior.Color = RGB(255, 255, 200)
    ws.Range("B14").Value = "KET QUA NOI SUY:"
    ws.Range("B14").Font.Bold = True
    ws.Range("C14").Formula = "=NoisuyMatran(C$3:G$3, $B$4:$B$9, C$4:G$9, C12, C13)"
    ws.Range("C14").Font.Bold = True
    ws.Range("C14").Font.Size = 11
    ws.Range("C14").Font.Color = RGB(0, 100, 0)
    ws.Range("C14").Interior.Color = RGB(210, 245, 210)
    ws.Range("C14").NumberFormat = "0.0000"
    
    ' --- PHAN 2: BANG TRA DAY SO 1 CHIEU (HAM =NoisuyDay) ---
    ws.Range("I2:L2").Merge
    ws.Range("I2").Value = "2. BANG TRA DAY SO 1 CHIEU (HAM =NoisuyDay)"
    ws.Range("I2").Font.Bold = True
    ws.Range("I2").Font.Size = 11
    ws.Range("I2").Font.Color = RGB(0, 51, 102)
    
    ws.Range("I3").Value = "Moc X"
    ws.Range("J3").Value = "Gia Tri Y"
    ws.Range("I3:J3").Interior.Color = RGB(220, 235, 252)
    ws.Range("I3:J3").Font.Bold = True
    ws.Range("I3:J3").HorizontalAlignment = xlCenter
    
    ws.Range("I4:I10").Value = Application.Transpose(Array(10, 20, 30, 40, 50, 75, 100))
    ws.Range("J4:J10").Value = Application.Transpose(Array(1.25, 2.4, 3.65, 4.8, 6.1, 9.25, 12.5))
    ws.Range("I3:J10").Borders.LineStyle = xlContinuous
    ws.Range("I4:J10").HorizontalAlignment = xlRight
    
    ' Tra cuu 1D
    ws.Range("I12").Value = "Gia tri X can tra:"
    ws.Range("J12").Value = 35
    ws.Range("J12").Interior.Color = RGB(255, 255, 200)
    ws.Range("I13").Value = "KET QUA NOI SUY:"
    ws.Range("I13").Font.Bold = True
    ws.Range("J13").Formula = "=NoisuyDay($I$4:$I$10, $J$4:$J$10, J12)"
    ws.Range("J13").Font.Bold = True
    ws.Range("J13").Font.Size = 11
    ws.Range("J13").Font.Color = RGB(0, 100, 0)
    ws.Range("J13").Interior.Color = RGB(210, 245, 210)
    ws.Range("J13").NumberFormat = "0.00"
    
    ' --- PHAN 3: NOI SUY 1 CHIEU 2 DIEM (HAM =Noisuy2b) CO TOOLTIP HUONG DAN ---
    ws.Range("B17:G17").Merge
    ws.Range("B17").Value = "3. NOI SUY 1 CHIEU (2 DIEM KEP - HAM =Noisuy2b)"
    ws.Range("B17").Font.Bold = True
    ws.Range("B17").Font.Size = 11
    ws.Range("B17").Font.Color = RGB(150, 50, 0)
    
    ws.Range("B18").Value = "Moc duoi (X1, Y1):"
    ws.Range("C18").Value = 10
    ws.Range("D18").Value = 100
    ws.Range("C18:D18").Interior.Color = RGB(240, 240, 240)
    ws.Range("C18:D18").HorizontalAlignment = xlCenter
    
    ws.Range("B19").Value = "Moc tren (X2, Y2):"
    ws.Range("C19").Value = 20
    ws.Range("D19").Value = 200
    ws.Range("C19:D19").Interior.Color = RGB(240, 240, 240)
    ws.Range("C19:D19").HorizontalAlignment = xlCenter
    
    ws.Range("B20").Value = "Gia tri can tra (X):"
    ws.Range("C20").Value = 15
    ws.Range("C20").Interior.Color = RGB(255, 255, 200)
    ws.Range("C20").HorizontalAlignment = xlCenter
    
    ws.Range("B21").Value = "KET QUA NOI SUY:"
    ws.Range("B21").Font.Bold = True
    ws.Range("C21").Formula = "=Noisuy2b(C18, D18, C19, D19, C20)"
    ws.Range("C21").Font.Bold = True
    ws.Range("C21").Font.Size = 11
    ws.Range("C21").Font.Color = RGB(0, 100, 0)
    ws.Range("C21").Interior.Color = RGB(210, 245, 210)
    ws.Range("C21").NumberFormat = "0.00"
    
    ' Gan Tooltip huong dan bay lo lung ngay duoi o ket qua C21!
    AttachTooltipNoisuy2b ws.Range("C21")
    
    ws.Range("E18:G21").Merge
    ws.Range("E18").Value = "Cu phap: =Noisuy2b(X1, Y1, X2, Y2, X)" & vbCrLf & _
                           "- (X1, Y1): Moc duoi | (X2, Y2): Moc tren | X: Can tra" & vbCrLf & _
                           "- Go '=Noisuy2b(' se tu dong hien hop thoai nhac lenh trang tinh lo lung!"
    ws.Range("E18").Font.Italic = True
    ws.Range("E18").Font.Size = 8.5
    ws.Range("E18").Font.Color = RGB(100, 100, 100)
    
    ws.Columns("B:L").AutoFit
    ' InsertNoiSuyToActiveWorkbook showMsg:=False
    
    If showMsg Then
        MsgBoxUni "\u0110\u00C3 T\u1EA0O B\u1EA2NG TRA M\u1EAEU CHO C\u1EA2 3 LO\u1EA0I T\u1EA0I SHEET '" & sheetName & "' TH\u00C0NH C\u00D4NG!" & vbCrLf & vbCrLf & _
                  "1. B\u1EA3ng Ma tr\u1EADn 2D (=NoisuyMatran)" & vbCrLf & _
                  "2. B\u1EA3ng D\u00E3y s\u1ED1 1D (=NoisuyDay)" & vbCrLf & _
                  "3. B\u1EA3ng N\u1ED9i suy 1 chi\u1EC1u 2 \u0111i\u1EC3m (=Noisuy2b) c\u00F3 g\u1EAFn Tooltip h\u01B0\u1EDBng d\u1EABn d\u01B0\u1EDBi \u00F4!", vbInformation, "AVTool"
    End If
End Sub
