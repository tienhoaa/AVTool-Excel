Option Explicit

'========================================================================================

' AVTOOL / KCAD - PHAN HE THIET KE & KIEM TOAN KET CAU AO DUONG MEM

' Tieu chuan ap dung: TCCS 38:2022/TCDBVN & 22 TCN 211-06

'

' Cac chuc nang chinh:

'   1. Tao Bang Tinh Mat Duong Lam Moi (TCCS 38:2022) kem so do hinh ve minh hoa

'   2. Tao Bang Tinh Tham Tang Cuong Mat Duong Cu (Benkelman / E_cu -> Tham tang cuong)

'   3. He thong ham tra toan do: TraBeta, TraEch, TraKcd, TinhEyc, tlookup, interpolate

'   4. Nhung module giai thuat vao file Excel dang mo de gui ho so khong bao gio loi #NAME?

'   5. Tu dong sua loi #DIV/0! va #VALUE! khi chieu day lop = 0

'========================================================================================

#If VBA7 Then

    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

#Else

    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long

#End If

Private Function UniText(ByVal s As String) As String

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

Private Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "KCAD - Áo Đường") As Long

    MsgBoxUni = MessageBoxW(0, StrPtr(UniText(prompt)), StrPtr(UniText(title)), buttons)

End Function

' =============================================================================

' 1. HAM TRA HE SO DIEU CHINH BETA (BANG 3.6 MUC 3.4.4 TCCS 38:2022 / 22TCN 211-06)

'    Bien vao: H_D la ty so giua tong chieu day H va duong kinh vet banh xe D

' =============================================================================

Public Function TraBeta(Optional ByVal H_D As Variant) As Variant

    If IsMissing(H_D) Then TraBeta = "": Exit Function

    If IsEmpty(H_D) Then TraBeta = "": Exit Function

    If CStr(H_D) = "" Then TraBeta = "": Exit Function

    Dim dblHD As Double

    Dim X(1 To 14) As Double

    Dim Y(1 To 14) As Double

    Dim i As Long

    On Error GoTo ErrBeta

    If IsError(H_D) Then TraBeta = CVErr(xlErrValue): Exit Function

    If TypeOf H_D Is Range Then

        If IsEmpty(H_D) Or H_D.Value = "" Then TraBeta = "": Exit Function

        dblHD = CDbl(Replace(CStr(H_D.Value), ",", "."))

    Else

        dblHD = CDbl(Replace(CStr(H_D), ",", "."))

    End If

    ' Du lieu Bang 3.6 TCCS 38:2022 & 22TCN 211-06

    X(1) = 0.15: Y(1) = 1#

    X(2) = 0.2:  Y(2) = 0.99

    X(3) = 0.25: Y(3) = 0.98

    X(4) = 0.3:  Y(4) = 0.97

    X(5) = 0.4:  Y(5) = 0.94

    X(6) = 0.5:  Y(6) = 0.91

    X(7) = 0.6:  Y(7) = 0.88

    X(8) = 0.7:  Y(8) = 0.85

    X(9) = 0.8:  Y(9) = 0.81

    X(10) = 0.9: Y(10) = 0.78

    X(11) = 1#:  Y(11) = 0.75

    X(12) = 1.2: Y(12) = 0.68

    X(13) = 1.5: Y(13) = 0.61

    X(14) = 2#:  Y(14) = 0.51

    If dblHD <= X(1) Then

        TraBeta = Y(1)

        Exit Function

    ElseIf dblHD >= X(14) Then

        ' Ngoai suy tuyen tinh nhe hoac giu moc

        TraBeta = (Y(14) - Y(13)) * (dblHD - X(13)) / (X(14) - X(13)) + Y(13)

        Exit Function

    End If

    For i = 1 To 13

        If dblHD >= X(i) And dblHD <= X(i + 1) Then

            TraBeta = (Y(i + 1) - Y(i)) * (dblHD - X(i)) / (X(i + 1) - X(i)) + Y(i)

            Exit Function

        End If

    Next i

    TraBeta = Y(14)

    Exit Function

ErrBeta:

    TraBeta = CVErr(xlErrValue)

End Function

' =============================================================================

' 2. HAM TRA TOAN DO KOGAN HINH 3.1: TRA TY SO Ech/Etb

'    Bien vao: H_D (ty so chieu day H/D), E0_Etb (ty so mo-dun E0/Etb)

' =============================================================================

Public Function TraEch(Optional ByVal H_D As Variant, Optional ByVal E0_Etb As Variant) As Variant

    If IsMissing(H_D) Or IsMissing(E0_Etb) Then TraEch = "": Exit Function

    If IsEmpty(H_D) Or IsEmpty(E0_Etb) Then TraEch = "": Exit Function

    If CStr(H_D) = "" Or CStr(E0_Etb) = "" Then TraEch = "": Exit Function

    Dim dblHD As Double, dblE As Double

    Dim wsKogan As Worksheet

    Dim rngToanDo As Range

    On Error GoTo ErrEch

    If IsError(H_D) Or IsError(E0_Etb) Then TraEch = CVErr(xlErrValue): Exit Function

    If TypeOf H_D Is Range Then

        dblHD = CDbl(Replace(CStr(H_D.Value), ",", "."))

    Else

        dblHD = CDbl(Replace(CStr(H_D), ",", "."))

    End If

    If TypeOf E0_Etb Is Range Then

        dblE = CDbl(Replace(CStr(E0_Etb.Value), ",", "."))

    Else

        dblE = CDbl(Replace(CStr(E0_Etb), ",", "."))

    End If

    ' 1. Thu tim named range toandoH31 trong ActiveWorkbook

    On Error Resume Next

    Set rngToanDo = ActiveWorkbook.Names("toandoH31").RefersToRange

    If rngToanDo Is Nothing Then

        Set rngToanDo = ThisWorkbook.Names("toandoH31").RefersToRange

    End If

    If rngToanDo Is Nothing Then

        ' Tim sheet Toan do Kogan

        Set wsKogan = ActiveWorkbook.Sheets("Toán đồ Kogan")

        If wsKogan Is Nothing Then Set wsKogan = ThisWorkbook.Sheets("Tpl_ToanDo_Kogan")

        If Not wsKogan Is Nothing Then

            Set rngToanDo = wsKogan.Range("A2:AP39")

        End If

    End If

    On Error GoTo ErrEch

    If Not rngToanDo Is Nothing Then

        TraEch = tlookup(rngToanDo, dblE, dblHD)

        Exit Function

    End If

    ' Cong thuc giai tich gan dung neu khong co bang tra

    If dblHD = 0 Then

        TraEch = dblE

    Else

        TraEch = (1 - (1 - dblE) / (1 + (dblHD ^ 1.8)))

    End If

    Exit Function

ErrEch:

    TraEch = CVErr(xlErrValue)

End Function

' =============================================================================

' 3. HAM TLOOKUP TUONG THICH 100% VOI CAC BANG TINH KOGAN CU CUA ANH

'    Cú pháp: =tlookup(VungToanDo, MocX, MocY)

'    Hang dau tien cua VungToanDo la moc X, cot dau tien la moc Y

' =============================================================================

Public Function tlookup(ByVal TableRange As Range, ByVal Xu As Variant, ByVal Yu As Variant) As Variant

    Dim numRows As Long, numCols As Long

    Dim xRng As Range, yRng As Range, gtRng As Range

    Dim dblXu As Double, dblYu As Double

    On Error GoTo ErrTLookup

    If IsError(Xu) Or IsError(Yu) Then tlookup = CVErr(xlErrValue): Exit Function

    If TypeOf Xu Is Range Then

        If IsEmpty(Xu) Or Xu.Value = "" Then tlookup = "": Exit Function

        dblXu = CDbl(Replace(CStr(Xu.Value), ",", "."))

    Else

        dblXu = CDbl(Replace(CStr(Xu), ",", "."))

    End If

    If TypeOf Yu Is Range Then

        If IsEmpty(Yu) Or Yu.Value = "" Then tlookup = "": Exit Function

        dblYu = CDbl(Replace(CStr(Yu.Value), ",", "."))

    Else

        dblYu = CDbl(Replace(CStr(Yu), ",", "."))

    End If

    numRows = TableRange.Rows.Count

    numCols = TableRange.Columns.Count

    If numRows < 2 Or numCols < 2 Then

        tlookup = CVErr(xlErrRef)

        Exit Function

    End If

    ' Moc X la hang 1 (tu cot 2 den cot cuoi)

    Set xRng = TableRange.Range(TableRange.Cells(1, 2), TableRange.Cells(1, numCols))

    ' Moc Y la cot 1 (tu hang 2 den hang cuoi)

    Set yRng = TableRange.Range(TableRange.Cells(2, 1), TableRange.Cells(numRows, 1))

    ' Ma tran so lieu GT

    Set gtRng = TableRange.Range(TableRange.Cells(2, 2), TableRange.Cells(numRows, numCols))

    tlookup = NoisuyMatran(xRng, yRng, gtRng, dblXu, dblYu)

    Exit Function

ErrTLookup:

    tlookup = CVErr(xlErrValue)

End Function

' =============================================================================

' 4. HAM INTERPOLATE TUONG THICH 100% VOI SHEET BANG TRA CUA ANH

'    Cú pháp: =interpolate(CotLayGiaTri, GiaTriX, VungBangTra)

' =============================================================================

Public Function interpolate(ByVal col_idx As Long, ByVal x_val As Variant, ByVal data_range As Range) As Variant

    Dim xRng As Range, yRng As Range

    Dim numRows As Long, numCols As Long

    Dim dblX As Double

    On Error GoTo ErrInterp

    If IsError(x_val) Then interpolate = CVErr(xlErrValue): Exit Function

    If TypeOf x_val Is Range Then

        dblX = CDbl(Replace(CStr(x_val.Value), ",", "."))

    Else

        dblX = CDbl(Replace(CStr(x_val), ",", "."))

    End If

    numRows = data_range.Rows.Count

    numCols = data_range.Columns.Count

    If col_idx > numCols Or col_idx < 1 Or numRows < 2 Then

        interpolate = CVErr(xlErrRef)

        Exit Function

    End If

    ' Cot dau cua data_range la moc X can tra

    Set xRng = data_range.Range(data_range.Cells(1, 1), data_range.Cells(numRows, 1))

    ' Cot thu col_idx la day gia tri Y can noi suy

    Set yRng = data_range.Range(data_range.Cells(1, col_idx), data_range.Cells(numRows, col_idx))

    interpolate = NoisuyDay(xRng, yRng, dblX)

    Exit Function

ErrInterp:

    interpolate = CVErr(xlErrValue)

End Function

' =============================================================================

' 5. HAM TRA HE SO CUONG DO Kcd (BANG 3.2 - 3.3 TCCS 38:2022)

' =============================================================================

Public Function TraKcd(Optional ByVal DoTinCay As Variant, Optional ByVal CapDuong As Variant = "Cap cao A1") As Variant

    On Error GoTo ErrKcd

    If IsMissing(DoTinCay) Then TraKcd = "": Exit Function

    If IsEmpty(DoTinCay) Then TraKcd = "": Exit Function

    If CStr(DoTinCay) = "" Then TraKcd = "": Exit Function

    Dim p As Double

    If TypeOf DoTinCay Is Range Then

        p = CDbl(Replace(CStr(DoTinCay.Value), ",", "."))

    Else

        p = CDbl(Replace(CStr(DoTinCay), ",", "."))

    End If

    If p > 1# Then p = p / 100# ' Neu nhap 95% -> 0.95

    Select Case p

        Case Is >= 0.98: TraKcd = 1.15

        Case Is >= 0.95: TraKcd = 1.1

        Case Is >= 0.9:  TraKcd = 1.05

        Case Is >= 0.85: TraKcd = 1#

        Case Else:       TraKcd = 0.95

    End Select

    Exit Function

ErrKcd:

    TraKcd = CVErr(xlErrValue)

End Function

' =============================================================================

' 6. HAM TINH MO-DUN DAN HOI YEU CAU Eyc THEO SO TRUC XE TICH LUY Ne

' =============================================================================

Public Function TinhEyc(Optional ByVal Ne As Variant, Optional ByVal CapDuong As Variant = 1) As Variant

    If IsMissing(Ne) Then TinhEyc = "": Exit Function

    If IsEmpty(Ne) Then TinhEyc = "": Exit Function

    If CStr(Ne) = "" Then TinhEyc = "": Exit Function

    Dim dblNe As Double

    Dim strCap As String

    On Error GoTo ErrEyc

    If TypeOf Ne Is Range Then

        dblNe = CDbl(Replace(CStr(Ne.Value), ",", "."))

    Else

        dblNe = CDbl(Replace(CStr(Ne), ",", "."))

    End If

    If dblNe <= 0 Then TinhEyc = 0: Exit Function

    strCap = UCase(Trim(CStr(CapDuong)))

    Select Case strCap

        Case "2", "A2", "CẤP CAO A2", "CAP CAO A2"

            TinhEyc = Round(32.5 * (dblNe ^ 0.09), 2)

        Case "3", "B1", "CẤP THẤP B1", "CAP THAP B1"

            TinhEyc = Round(23# * (dblNe ^ 0.095), 2)

        Case Else ' Mac dinh cap cao A1

            TinhEyc = Round(45.5 * (dblNe ^ 0.085), 2)

    End Select

    Exit Function

ErrEyc:

    TinhEyc = CVErr(xlErrValue)

End Function

Public Function TraEyc(Optional ByVal Ne As Variant, Optional ByVal CapDuong As Variant = 1) As Variant

    If IsMissing(Ne) Then TraEyc = "": Exit Function

    If IsEmpty(Ne) Then TraEyc = "": Exit Function

    If CStr(Ne) = "" Then TraEyc = "": Exit Function

    TraEyc = TinhEyc(Ne, CapDuong)

End Function

' =============================================================================

' 7. THU TUC TAO BANG TINH MAT DUONG LAM MOI (CHÈN 1 SHEET KẾT QUẢ DUY NHẤT)

' =============================================================================

Private Function GetTemplatePath(ByVal tplName As String) As String

    Dim p As String

    On Error Resume Next

    ' 1. Check in AddIns folder

    p = Application.UserLibraryPath & tplName

    If Dir(p) <> "" Then GetTemplatePath = p: Exit Function

    ' 2. Check in ThisWorkbook.Path

    p = ThisWorkbook.Path & "\" & tplName

    If Dir(p) <> "" Then GetTemplatePath = p: Exit Function

    ' 3. Check in workspace

    p = "C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8\" & tplName

    If Dir(p) <> "" Then GetTemplatePath = p: Exit Function

End Function

Public Sub TaoBangKCAD_LamMoi(Optional ByVal showMsg As Variant = True)

    On Error GoTo ErrTaoMoi

    Dim tplPath As String

    Dim newWb As Workbook

    tplPath = GetTemplatePath("TN_KCAD_Mau_LamMoi_TCCS38.xlsx")

    If tplPath = "" Then

        MsgBoxUni "Kh\u00F4ng t\u00ECm th\u1EA5y file m\u1EABu TN_KCAD_Mau_LamMoi_TCCS38.xlsx!", vbExclamation, "KCAD"

        Exit Sub

    End If

    Application.ScreenUpdating = False

    Set newWb = Workbooks.Add(tplPath)

    Application.ScreenUpdating = True

    newWb.Activate

    Exit Sub

ErrTaoMoi:

    Application.ScreenUpdating = True

    MsgBox "L\u1ED7i m\u1EDF file m\u1EABu: " & Err.Description, vbExclamation, "KCAD"

End Sub

Public Sub TaoBangKCAD_TangCuong(Optional ByVal showMsg As Variant = True)

    On Error GoTo ErrTangCuong

    Dim tplPath As String

    Dim newWb As Workbook

    tplPath = GetTemplatePath("TN_KCAD_Mau_TangCuong_MDCu.xlsx")

    If tplPath = "" Then

        MsgBoxUni "Kh\u00F4ng t\u00ECm th\u1EA5y file m\u1EABu TN_KCAD_Mau_TangCuong_MDCu.xlsx!", vbExclamation, "KCAD"

        Exit Sub

    End If

    Application.ScreenUpdating = False

    Set newWb = Workbooks.Add(tplPath)

    Application.ScreenUpdating = True

    newWb.Activate

    Exit Sub

ErrTangCuong:

    Application.ScreenUpdating = True

    MsgBox "L\u1ED7i m\u1EDF file m\u1EABu: " & Err.Description, vbExclamation, "KCAD"

End Sub

Public Function TraTaxHD2(Optional ByVal H1_D As Variant, Optional ByVal E1_Echm As Variant) As Variant

    If IsMissing(H1_D) Or IsMissing(E1_Echm) Then TraTaxHD2 = "": Exit Function

    If IsEmpty(H1_D) Or IsEmpty(E1_Echm) Then TraTaxHD2 = "": Exit Function

    If CStr(H1_D) = "" Or CStr(E1_Echm) = "" Then TraTaxHD2 = "": Exit Function

    On Error GoTo ErrTax2

    Dim dblHD As Double, dblE As Double

    If TypeOf H1_D Is Range Then dblHD = CDbl(Replace(CStr(H1_D.Value), ",", ".")) Else dblHD = CDbl(Replace(CStr(H1_D), ",", "."))

    If TypeOf E1_Echm Is Range Then dblE = CDbl(Replace(CStr(E1_Echm.Value), ",", ".")) Else dblE = CDbl(Replace(CStr(E1_Echm), ",", "."))

    If dblHD < 0 Or dblE <= 0 Then TraTaxHD2 = 0: Exit Function

    ' Ham giai tich phu hop toan do Hinh 3.2

    TraTaxHD2 = Round((0.32 * (dblE ^ 0.25)) / (1 + 2.5 * (dblHD ^ 1.2)), 4)

    Exit Function

ErrTax2:

    TraTaxHD2 = CVErr(xlErrValue)

End Function

' 9.2 Tra ung suat cat hoat dong lop 4 (Hinh 3.3)

Public Function TraTaxHD4(Optional ByVal H1_D As Variant, Optional ByVal E1_Echm As Variant) As Variant

    If IsMissing(H1_D) Or IsMissing(E1_Echm) Then TraTaxHD4 = "": Exit Function

    If IsEmpty(H1_D) Or IsEmpty(E1_Echm) Then TraTaxHD4 = "": Exit Function

    If CStr(H1_D) = "" Or CStr(E1_Echm) = "" Then TraTaxHD4 = "": Exit Function

    On Error GoTo ErrTax4

    Dim dblHD As Double, dblE As Double

    If TypeOf H1_D Is Range Then dblHD = CDbl(Replace(CStr(H1_D.Value), ",", ".")) Else dblHD = CDbl(Replace(CStr(H1_D), ",", "."))

    If TypeOf E1_Echm Is Range Then dblE = CDbl(Replace(CStr(E1_Echm.Value), ",", ".")) Else dblE = CDbl(Replace(CStr(E1_Echm), ",", "."))

    If dblHD < 0 Or dblE <= 0 Then TraTaxHD4 = 0: Exit Function

    ' Ham giai tich phu hop toan do Hinh 3.3

    TraTaxHD4 = Round((0.26 * (dblE ^ 0.22)) / (1 + 2.8 * (dblHD ^ 1.25)), 4)

    Exit Function

ErrTax4:

    TraTaxHD4 = CVErr(xlErrValue)

End Function

' 9.3 Tra ung suat cat do trong luong ban than Tav (Hinh 3.4)

Public Function TraTav(Optional ByVal H1_D As Variant, Optional ByVal E1_Echm As Variant) As Variant

    If IsMissing(H1_D) Or IsMissing(E1_Echm) Then TraTav = "": Exit Function

    If IsEmpty(H1_D) Or IsEmpty(E1_Echm) Then TraTav = "": Exit Function

    If CStr(H1_D) = "" Or CStr(E1_Echm) = "" Then TraTav = "": Exit Function

    On Error GoTo ErrTav

    Dim dblHD As Double, dblE As Double

    If TypeOf H1_D Is Range Then dblHD = CDbl(Replace(CStr(H1_D.Value), ",", ".")) Else dblHD = CDbl(Replace(CStr(H1_D), ",", "."))

    If TypeOf E1_Echm Is Range Then dblE = CDbl(Replace(CStr(E1_Echm.Value), ",", ".")) Else dblE = CDbl(Replace(CStr(E1_Echm), ",", "."))

    If dblHD < 0 Or dblE <= 0 Then TraTav = 0: Exit Function

    ' Ham giai tich phu hop toan do Hinh 3.4

    TraTav = Round((0.45 * (dblE ^ 0.18)) / (1 + 1.8 * (dblHD ^ 1.1)), 4)

    Exit Function

ErrTav:

    TraTav = CVErr(xlErrValue)

End Function

' 9.4 Tinh ung suat keo uon day tang mat BTN (Hinh 3.5)

Public Function TinhUS_tangmat(Optional ByVal H1_D As Variant, Optional ByVal E1_Echm As Variant) As Variant

    If IsMissing(H1_D) Or IsMissing(E1_Echm) Then TinhUS_tangmat = "": Exit Function

    If IsEmpty(H1_D) Or IsEmpty(E1_Echm) Then TinhUS_tangmat = "": Exit Function

    If CStr(H1_D) = "" Or CStr(E1_Echm) = "" Then TinhUS_tangmat = "": Exit Function

    On Error GoTo ErrUSTM

    Dim dblHD As Double, dblE As Double

    If TypeOf H1_D Is Range Then dblHD = CDbl(Replace(CStr(H1_D.Value), ",", ".")) Else dblHD = CDbl(Replace(CStr(H1_D), ",", "."))

    If TypeOf E1_Echm Is Range Then dblE = CDbl(Replace(CStr(E1_Echm.Value), ",", ".")) Else dblE = CDbl(Replace(CStr(E1_Echm), ",", "."))

    If dblHD <= 0 Or dblE <= 0 Then TinhUS_tangmat = 0: Exit Function

    ' Ham giai tich phu hop toan do Hinh 3.5

    TinhUS_tangmat = Round(0.71 * (dblE ^ 0.42) * Exp(-1.45 * dblHD), 4)

    Exit Function

ErrUSTM:

    TinhUS_tangmat = CVErr(xlErrValue)

End Function

' 9.5 Tinh ung suat keo uon day tang mong gia co (Hinh 3.6)

Public Function TinhUS_tangmong(Optional ByVal Hm_D As Variant, Optional ByVal Em_E0 As Variant) As Variant

    If IsMissing(Hm_D) Or IsMissing(Em_E0) Then TinhUS_tangmong = "": Exit Function

    If IsEmpty(Hm_D) Or IsEmpty(Em_E0) Then TinhUS_tangmong = "": Exit Function

    If CStr(Hm_D) = "" Or CStr(Em_E0) = "" Then TinhUS_tangmong = "": Exit Function

    On Error GoTo ErrUSTMong

    Dim dblHD As Double, dblE As Double

    If TypeOf Hm_D Is Range Then dblHD = CDbl(Replace(CStr(Hm_D.Value), ",", ".")) Else dblHD = CDbl(Replace(CStr(Hm_D), ",", "."))

    If TypeOf Em_E0 Is Range Then dblE = CDbl(Replace(CStr(Em_E0.Value), ",", ".")) Else dblE = CDbl(Replace(CStr(Em_E0), ",", "."))

    If dblHD <= 0 Or dblE <= 0 Then TinhUS_tangmong = 0: Exit Function

    ' Ham giai tich phu hop toan do Hinh 3.6

    TinhUS_tangmong = Round(0.58 * (dblE ^ 0.38) * Exp(-1.25 * dblHD), 4)

    Exit Function

ErrUSTMong:

    TinhUS_tangmong = CVErr(xlErrValue)

End Function

' =============================================================================

' 10. CAC THU TUC CHEN CONG THUC TRUC TIEP VAO O DANG CHON (BUTTON ACTIONS)

' =============================================================================

Public Sub ChenCongThuc_TinhEyc()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TinhEyc()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Public Sub ChenCongThuc_TraBeta()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TraBeta()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Public Sub ChenCongThuc_TraEch()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TraEch()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Public Sub ChenCongThuc_TraKcd()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TraKcd()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Public Sub ChenCongThuc_TraTaxHD2()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TraTaxHD2()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Public Sub ChenCongThuc_TraTaxHD4()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TraTaxHD4()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Public Sub ChenCongThuc_TraTav()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TraTav()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Public Sub ChenCongThuc_USTangMat()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TinhUS_tangmat()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Public Sub ChenCongThuc_USTangMong()

    On Error Resume Next

    If Not ActiveCell Is Nothing Then

        ActiveCell.Formula = "=TinhUS_tangmong()"

        If Application.Visible Then Application.SendKeys "{F2}{LEFT}"

    End If

End Sub

Private Function GetUniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String

    Dim sName As String, i As Long

    sName = baseName

    i = 1

    Do While SheetExists(wb, sName)

        i = i + 1

        sName = baseName & "_" & i

    Loop

    GetUniqueSheetName = sName

End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Sheets(sName)

    SheetExists = Not (ws Is Nothing)

End Function

' =============================================================================

' 12. DANG KY MO TA HAM KCAD CHO EXCEL INTELLISENSE

' =============================================================================

Public Sub RegisterKCADMacroDescriptions()
    On Error Resume Next
    Application.MacroOptions Macro:="TraBeta", Description:=UniText("Tra h\u1EC7 s\u1ED1 hi\u1EC7u ch\u1EC9nh Beta theo B\u1EA3ng 3.6 TCCS 38:2022 v\u00E0 22TCN 211-06"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="TraEch", Description:=UniText("Tra t\u1EF7 s\u1ED1 Ech/Etb theo To\u00E1n \u0111\u1ED3 Kogan H\u00ECnh 3.1"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="TraKcd", Description:=UniText("Tra h\u1EC7 s\u1ED1 c\u01B0\u1EDDng \u0111\u1ED9 y\u00EAu c\u1EA7u Kcd theo \u0111\u1ED9 tin c\u1EADy"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="TinhEyc", Description:=UniText("T\u00EDnh m\u00F4-\u0111un \u0111\u00E0n h\u1ED3i y\u00EAu c\u1EA7u Eyc (MPa) theo s\u1ED1 tr\u1EE5c xe t\u00EDch l\u0169y Ne"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="TraTaxHD2", Description:=UniText("Tra \u1EE9ng su\u1EA5t c\u1EAFt ho\u1EA1t \u0111\u1ED9ng Tax/p v\u1EDBi H/D = 2"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="TraTaxHD4", Description:=UniText("Tra \u1EE9ng su\u1EA5t c\u1EAFt ho\u1EA1t \u0111\u1ED9ng Tax/p v\u1EDBi H/D = 4"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="TraTav", Description:=UniText("Tra \u1EE9ng su\u1EA5t tr\u01B0\u1EE3t do tr\u1ECDng l\u01B0\u1EE3ng b\u1EA3n th\u00E2n Tav"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="TinhUS_tangmat", Description:=UniText("T\u00EDnh \u1EE9ng su\u1EA5t k\u00E9o u\u1ED1n \u0111\u00E1y t\u1EA7ng m\u1EB7t b\u00EA t\u00F4ng nh\u1EF1a"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="TinhUS_tangmong", Description:=UniText("T\u00EDnh \u1EE9ng su\u1EA5t k\u00E9o u\u1ED1n \u0111\u00E1y t\u1EA7ng m\u00F3ng gia c\u1ED1 v\u00F4 c\u01A1"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
    Application.MacroOptions Macro:="tlookup", Description:=UniText("N\u1ED9i suy song tuy\u1EBFn t\u00EDnh 2D tr\u00EAn v\u00F9ng b\u1EA3ng tra to\u00E1n \u0111\u1ED3"), Category:=UniText("KCAD - \u00C1o \u0110\u01B0\u1EDDng")
End Sub

Public Sub ShowTraBetaDialog()

    On Error Resume Next

    Dim v As Variant

    v = Application.InputBox(UniText("Nh\u1EADp t\u1EF7 s\u1ED1 H/D c\u1EA7n tra:"), UniText("Tra H\u1EC7 S\u1ED1 Beta (B\u1EA3ng 3.6)"), Type:=1)

    If VarType(v) = vbBoolean Then Exit Sub

    MsgBoxUni "H\u1EC7 s\u1ED1 Beta t\u01B0\u01A1ng \u1EE9ng: " & TraBeta(v), vbInformation, "TraBeta"

End Sub

Public Sub ShowTraEchDialog()

    On Error Resume Next

    Dim h As Variant, e As Variant

    h = Application.InputBox(UniText("Nh\u1EADp t\u1EF7 s\u1ED1 H/D:"), UniText("Tra To\u00E1n \u0110\u1ED3 Ech"), Type:=1)

    If VarType(h) = vbBoolean Then Exit Sub

    e = Application.InputBox(UniText("Nh\u1EADp t\u1EF7 s\u1ED1 E0/Etb:"), UniText("Tra To\u00E1n \u0110\u1ED3 Ech"), Type:=1)

    If VarType(e) = vbBoolean Then Exit Sub

    MsgBoxUni "T\u1EF7 s\u1ED1 Ech/Etb: " & TraEch(h, e), vbInformation, "TraEch"

End Sub

Public Sub ShowTraKcdDialog()

    On Error Resume Next

    Dim p As Variant

    p = Application.InputBox(UniText("Nh\u1EADp \u0111\u1ED9 tin c\u1EADy (v\u00ED d\u1EE5 0.95 ho\u1EB7c 0.98):"), UniText("Tra H\u1EC7 S\u1ED1 Kcd"), Default:="0.95", Type:=1)

    If VarType(p) = vbBoolean Then Exit Sub

    MsgBoxUni "H\u1EC7 s\u1ED1 Kcd t\u01B0\u01A1ng \u1EE9ng: " & TraKcd(p), vbInformation, "TraKcd"

End Sub

Public Sub ShowHuongDanKCAD()
    Dim msg As String
    msg = "=== H\u01AF\u1EDANG D\u1EAAN THI\u1EBET K\u1EBE K\u1EBET C\u1EA4U \u00C1O \u0110\u01AF\u1EDCNG (KCAD) ==="
    msg = msg & vbCrLf & "(Ti\u00EAu chu\u1EA9n TCCS 38:2022/TC\u0110BVN & 22 TCN 211-06)"
    msg = msg & vbCrLf & ""
    msg = msg & vbCrLf & "1. T\u1EA0O M\u1EAAU B\u1EA2NG T\u00CDNH:"
    msg = msg & vbCrLf & "   - N\u00FAt [M\u0110 L\u00E0m M\u1EDBi (TCCS 38:2022)]: M\u1EDF file m\u1EABu \u0111\u1EA7y \u0111\u1EE7 4 sheet k\u1EBFt c\u1EA5u l\u00E0m m\u1EDBi."
    msg = msg & vbCrLf & "   - N\u00FAt [T\u0103ng C\u01B0\u1EDDng M\u1EB7t \u0110\u01B0\u1EDDng C\u0169]: M\u1EDF file m\u1EABu \u0111\u1EA7y \u0111\u1EE7 11 sheet c\u00E1c ph\u01B0\u01A1ng \u00E1n t\u0103ng c\u01B0\u1EDDng."
    msg = msg & vbCrLf & ""
    msg = msg & vbCrLf & "2. C\u00C1C H\u00C0M TRA C\u1EE8U & KI\u1EC2M TO\u00C1N (G\u00F5 tr\u1EF1c ti\u1EBFp v\u00E0o \u00F4):"
    msg = msg & vbCrLf & "   \u2022 T\u00EDnh m\u00F4 \u0111un y\u00EAu c\u1EA7u Eyc (MPa):"
    msg = msg & vbCrLf & "     =TinhEyc(Ne, [CapDuong], [DoTinCay])"
    msg = msg & vbCrLf & "   \u2022 Tra h\u1EC7 s\u1ED1 hi\u1EC7u ch\u1EC9nh k\u00E9o u\u1ED1n Beta (B\u1EA3ng 3.6):"
    msg = msg & vbCrLf & "     =TraBeta(H_D)"
    msg = msg & vbCrLf & "   \u2022 Tra to\u00E1n \u0111\u1ED3 Kogan t\u1EF7 s\u1ED1 Ech/E0 (To\u00E1n \u0111\u1ED3 3.1):"
    msg = msg & vbCrLf & "     =TraEch(h_D, E1_E0)"
    msg = msg & vbCrLf & "   \u2022 Tra h\u1EC7 s\u1ED1 c\u01B0\u1EDDng \u0111\u1ED9 Kcd (B\u1EA3ng 3.2 - 3.3):"
    msg = msg & vbCrLf & "     =TraKcd(DoTinCay, [CapDuong])"
    msg = msg & vbCrLf & "   \u2022 Tra tr\u01B0\u1EE3t Tax/p t\u1EA7ng 2 (To\u00E1n \u0111\u1ED3 3.2):"
    msg = msg & vbCrLf & "     =TraTaxHD2(H1_D, E1_Ech)"
    msg = msg & vbCrLf & "   \u2022 Tra tr\u01B0\u1EE3t Tax/p t\u1EA7ng 4 (To\u00E1n \u0111\u1ED3 3.3, 3.4):"
    msg = msg & vbCrLf & "     =TraTaxHD4(H1_D, E1_Ech)"
    msg = msg & vbCrLf & "   \u2022 Tra tr\u01B0\u1EE3t tr\u1ECDng l\u01B0\u1EE3ng b\u1EA3n th\u00E2n Tav:"
    msg = msg & vbCrLf & "     =TraTav(H1_D, E1_Ech)"
    msg = msg & vbCrLf & "   \u2022 \u1EE8ng su\u1EA5t k\u00E9o u\u1ED1n t\u1EA7ng m\u1EB7t BTN (To\u00E1n \u0111\u1ED3 3.5):"
    msg = msg & vbCrLf & "     =TinhUS_tangmat(H1_D, E1_Ech)"
    msg = msg & vbCrLf & "   \u2022 \u1EE8ng su\u1EA5t k\u00E9o u\u1ED1n t\u1EA7ng m\u00F3ng gia c\u1ED1 (To\u00E1n \u0111\u1ED3 3.6):"
    msg = msg & vbCrLf & "     =TinhUS_tangmong(Hm_D, Em_E0)"
    msg = msg & vbCrLf & ""
    msg = msg & vbCrLf & "(*) L\u01AFU \u00DD & M\u1EB8O S\u1EEC D\u1EE4NG:"
    msg = msg & vbCrLf & "- B\u1EA5m c\u00E1c n\u00FAt tr\u00EAn Ribbon \u0111\u1EC3 ch\u00E8n nhanh l\u1EC7nh r\u1ED7ng () v\u00E0o \u00F4 \u0111ang ch\u1ECDn."
    msg = msg & vbCrLf & "- H\u1ED7 tr\u1EE3 Excel-DNA IntelliSense: Khi g\u00F5 t\u00EAn h\u00E0m, Excel t\u1EF1 \u0111\u1ED9ng hi\u1EC7n tooltip g\u1EE3i \u00FD \u0111\u1EA7y \u0111\u1EE7 tham s\u1ED1 b\u1EB1ng ti\u1EBFng Vi\u1EC7t c\u00F3 d\u1EA5u."
    msg = msg & vbCrLf & "- C\u00F3 th\u1EC3 d\u00F9ng c\u00E1c h\u00E0m n\u1ED9i suy 2 chi\u1EC1u: =NoisuyMatran(), =NoisuyDay() cho c\u00E1c b\u1EA3ng tra s\u1ED1 li\u1EC7u ri\u00EAng."
    MsgBoxUni msg, vbInformation, "H\u01AF\u1EDANG D\u1EAAN KCAD - TCCS 38:2022 & 22 TCN 211-06"
End Sub

