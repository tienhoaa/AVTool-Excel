Option Explicit

'========================================================================================
' AVTOOL - BO CONG CU NOI SUY BANG TRA VA TINH TOAN KET CAU CHUYEN NGHIEP
'
' Bao gom day du cac tinh nang:
'   1. NoisuyMatran(Xk, Yk, GT, Xu, Yu): Noi suy ma tran 2 chieu (Bilinear Interpolation)
'   2. NoisuyDay(Xk, Yk, Xu): Noi suy 1 chieu theo day so (Linear Interpolation)
'   3. Noisuy2b(Xs1, Ys1, Xs2, Ys2, Xs): Noi suy tuyen tinh giua 2 diem co so
'   4. Dt(d): Tinh dien tich tiet dien cot thep tron (d tinh bang mm -> m2)
'   5. Cac Form truc quan cho nguoi moi: frmNoiSuyMatran, frmNoiSuyDay, frmNoiSuy2b, frmTinhThep
'   6. Tu dong nhung module vao file Excel dang mo de tranh loi #NAME? khi sang may khac
'========================================================================================

' -----------------------------------------------------------------------------
' 1. HAM TINH DIEN TICH COT THEP TRON (d theo mm -> dien tich m2)
' -----------------------------------------------------------------------------
Public Function Dt(ByVal d As Double) As Double
    On Error Resume Next
    Dt = 3.14159265358979 * (d ^ 2) / 4000000#
End Function

' -----------------------------------------------------------------------------
' 2. HAM NOI SUY TUYEN TINH 2 DIEM (Xs1, Ys1) VA (Xs2, Ys2)
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
' 3. THUAT TOAN PHU TRO NOI SUY 1 CHIEU CO TU DONG SAP XEP
' -----------------------------------------------------------------------------
Public Function TinhtoanNs(ByVal Ncot As Long, ByRef X() As Double, ByRef Y() As Double, ByVal Xu As Double, Optional ByVal tb As String = "") As Double
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    Dim i As Long, j As Long
    Dim hasError As Boolean
    
    On Error Resume Next
    Ncot = UBound(X)
    
    ' Sap xep noi bot tang dan theo X, dong thoi hoan vi Y
    For i = 1 To Ncot - 1
        For j = i + 1 To Ncot
            If X(i) > X(j) Then
                X1 = X(i): X(i) = X(j): X(j) = X1
                Y1 = Y(i): Y(i) = Y(j): Y(j) = Y1
            End If
        Next j
    Next i
    
    If (X(Ncot) < Xu) Or (X(1) > Xu) Then
        hasError = True
    End If
    
    If hasError Then
        TinhtoanNs = -100000000000#
    Else
        ' Neu trung dung moc
        For i = 1 To Ncot
            If X(i) = Xu Then
                TinhtoanNs = Y(i)
                Exit Function
            End If
        Next i
        
        ' Noi suy tuyen tinh trong khoang
        For i = 1 To Ncot - 1
            If (X(i) < Xu) And (X(i + 1) > Xu) Then
                X1 = X(i): Y1 = Y(i)
                X2 = X(i + 1): Y2 = Y(i + 1)
                TinhtoanNs = (Y2 - Y1) * (Xu - X1) / (X2 - X1) + Y1
                Exit Function
            End If
        Next i
    End If
End Function

' -----------------------------------------------------------------------------
' 4. HAM NOI SUY 1 CHIEU THEO DAY SO (Xk: day moc X, Yk: day moc Y, Xu: gia tri can tra)
' -----------------------------------------------------------------------------
Public Function NoisuyDay(ByVal Xk As Range, ByVal Yk As Range, ByVal Xu As Double) As Double
    Dim X() As Double, Y() As Double
    Dim Ncot As Long, i As Long, j As Long, k As Long
    
    On Error GoTo LoiDay
    Ncot = Xk.Count
    If Yk.Count <> Ncot Then
        NoisuyDay = -100000000000#
        Exit Function
    End If
    
    ReDim X(Ncot)
    ReDim Y(Ncot)
    
    k = 0
    For i = 1 To Xk.Rows.Count
        For j = 1 To Xk.Columns.Count
            k = k + 1
            X(k) = CDbl(Xk.Item(i, j).Value)
        Next j
    Next i
    
    k = 0
    For i = 1 To Yk.Rows.Count
        For j = 1 To Yk.Columns.Count
            k = k + 1
            Y(k) = CDbl(Yk.Item(i, j).Value)
        Next j
    Next i
    
    NoisuyDay = TinhtoanNs(Ncot, X, Y, Xu, "")
    Exit Function
    
LoiDay:
    NoisuyDay = -100000000000#
End Function

' -----------------------------------------------------------------------------
' 5. HAM NOI SUY MA TRAN 2 CHIEU SONG TUYEN TINH (Bilinear Interpolation)
'    Xk: Hang moc ngang (1 dong x N cot)
'    Yk: Cot moc doc (M dong x 1 cot)
'    GT: Vung ma tran du lieu (M dong x N cot)
'    Xu: Gia tri tra theo truc X
'    Yu: Gia tri tra theo truc Y
' -----------------------------------------------------------------------------
Public Function NoisuyMatran(ByVal Xk As Range, ByVal Yk As Range, ByVal GT As Range, ByVal Xu As Double, ByVal Yu As Double) As Double
    Dim X() As Double, Y() As Double, G() As Double
    Dim i As Long, j As Long, k As Long
    Dim Ndong As Long, Ncot As Long
    Dim Phu As Double
    Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    Dim G1 As Double, G2 As Double, G3 As Double
    Dim Timthay As Boolean
    
    On Error GoTo LoiMatran
    Timthay = False
    Ndong = GT.Rows.Count
    Ncot = GT.Columns.Count
    
    If Xk.Count <> Ncot Or Yk.Count <> Ndong Then
        NoisuyMatran = 0
        Exit Function
    End If
    
    ReDim G(1 To Ndong, 1 To Ncot)
    For i = 1 To Ndong
        For j = 1 To Ncot
            G(i, j) = CDbl(GT.Item(i, j).Value)
        Next j
    Next i
    
    ReDim X(1 To Ncot)
    ReDim Y(1 To Ndong)
    
    For i = 1 To Ncot
        X(i) = CDbl(Xk.Item(1, i).Value)
    Next i
    For i = 1 To Ndong
        Y(i) = CDbl(Yk.Item(i, 1).Value)
    Next i
    
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
    
    ' Tim khoang bao chua Xu va Yu
    For i = 1 To Ndong - 1
        For j = 1 To Ncot - 1
            If (Y(i) <= Yu) And (Y(i + 1) >= Yu) And (X(j) <= Xu) And (X(j + 1) >= Xu) Then
                X1 = j: X2 = j + 1
                Y1 = i: Y2 = i + 1
                Timthay = True
                Exit For
            End If
        Next j
        If Timthay Then Exit For
    Next i
    
    If Timthay Then
        G1 = Noisuy2b(Y(Y1), G(Y1, X1), Y(Y2), G(Y2, X1), Yu)
        G2 = Noisuy2b(Y(Y1), G(Y1, X2), Y(Y2), G(Y2, X2), Yu)
        G3 = Noisuy2b(X(X1), G1, X(X2), G2, Xu)
        NoisuyMatran = G3
    Else
        NoisuyMatran = 0
    End If
    Exit Function

LoiMatran:
    NoisuyMatran = 0
End Function

' -----------------------------------------------------------------------------
' 6. HAM PHU TRO NAME (Tra ve ten dai o - Giu nguyen tu code goc)
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

' =============================================================================
' 7. DIEU KHIEN HIEN THI CAC FORM CHUYEN DUNG (MODELess DE THAO TAC TREN EXCEL)
' =============================================================================
Public Sub ShowNoiSuyMatran()
    On Error Resume Next
    If Val(Application.Version) >= 12 Then
        frmNoiSuyMatran.Show vbModeless
    Else
        frmNoiSuyMatran.Show
    End If
End Sub

Public Sub ShowNoiSuyDay()
    On Error Resume Next
    If Val(Application.Version) >= 12 Then
        frmNoiSuyDay.Show vbModeless
    Else
        frmNoiSuyDay.Show
    End If
End Sub

Public Sub ShowNoiSuy2b()
    On Error Resume Next
    If Val(Application.Version) >= 12 Then
        frmNoiSuy2b.Show vbModeless
    Else
        frmNoiSuy2b.Show
    End If
End Sub

Public Sub ShowTinhThep()
    On Error Resume Next
    If Val(Application.Version) >= 12 Then
        frmTinhThep.Show vbModeless
    Else
        frmTinhThep.Show
    End If
End Sub

' =============================================================================
' 8. TU DONG NHUNG MOD_NOISUY VAO FILE DANG MO (TRANH LOI #NAME? KHI GUI FILE)
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
        If showMsg Then MsgBox "Chua co file Excel nao dang mo!", vbExclamation, "AVTool"
        Exit Function
    End If
    
    If targetWb Is ThisWorkbook Then
        If showMsg Then MsgBox "Day chinh la file Add-in AVTool!", vbInformation, "AVTool"
        Exit Function
    End If
    
    Err.Clear
    Set vbProj = targetWb.VBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        If showMsg Then
            MsgBox "Khong the chen Module vao file vi Excel chua mo khoa quyen VBA." & vbCrLf & vbCrLf & _
                   "Cach bat rat nhanh (chi can lam 1 lan):" & vbCrLf & _
                   "1. Vao File > Options > Trust Center > Trust Center Settings..." & vbCrLf & _
                   "2. Chon Macro Settings > Tich chon 'Trust access to the VBA project object model' > OK.", vbExclamation, "AVTool"
        End If
        Exit Function
    End If
    
    Set srcComp = ThisWorkbook.VBProject.VBComponents("Mod_NoiSuy")
    If srcComp Is Nothing Then
        If showMsg Then MsgBox "Loi: Khong tim thay ma nguon Mod_NoiSuy trong Add-in!", vbCritical, "AVTool"
        Exit Function
    End If
    
    lineCount = srcComp.CodeModule.CountOfLines
    noiSuyCode = srcComp.CodeModule.Lines(1, lineCount)
    
    ' Kiem tra xem ham NoisuyMatran da ton tai trong bat ky module nao chua
    Set existingComp = Nothing
    For Each comp In vbProj.VBComponents
        If comp.Type = 1 Then ' vbext_ct_StdModule
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
        ' Da co san module, cap nhat code de co ban moi nhat
        existingComp.CodeModule.DeleteLines 1, existingComp.CodeModule.CountOfLines
        existingComp.CodeModule.AddFromString noiSuyCode
        InsertNoiSuyToActiveWorkbook = True
        If showMsg Then
            MsgBox "DA CAP NHAT BO HAM NOI SUY BAN MOI NHAT VAO FILE '" & targetWb.Name & "' THANH CONG!" & vbCrLf & vbCrLf & _
                   "Cac ham: NoisuyMatran, NoisuyDay, Noisuy2b, Dt da san sang trong Module '" & existingComp.Name & "'.", vbInformation, "AVTool"
        End If
    Else
        ' Tao moi Module Mod_NoiSuy
        Set comp = vbProj.VBComponents.Add(1)
        comp.Name = "Mod_NoiSuy"
        comp.CodeModule.AddFromString noiSuyCode
        InsertNoiSuyToActiveWorkbook = True
        If showMsg Then
            MsgBox "DA CHEN BO HAM NOI SUY VAO FILE '" & targetWb.Name & "' THANH CONG!" & vbCrLf & vbCrLf & _
                   "Module 'Mod_NoiSuy' da duoc nhung vinh vien vao file nay." & vbCrLf & _
                   "Luu y: Nho luu file duoi dang .xls hoac .xlsm de nguoi dung may khac khong can cai Add-in van chay duoc 100%!", vbInformation, "AVTool"
        End If
    End If
    
    On Error Resume Next
    If Not Application.VBE Is Nothing Then
        Application.VBE.MainWindow.Visible = False
    End If
    Application.Calculate
End Function

' =============================================================================
' 9. TAO BANG TRA MAU MINH HOA CHO NGUOI MOI DE DANG HIEU VA AP DUNG
' =============================================================================
Public Sub CreateSampleTableMatran(Optional ByVal showMsg As Boolean = True)
    Dim wb As Workbook, ws As Worksheet
    Dim sheetName As String
    Dim idx As Long
    
    On Error Resume Next
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Vui long mo mot file Excel de tao bang mau!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    ' Dat ten sheet mau
    sheetName = "Mau_NoiSuy_2D"
    idx = 1
    Do While True
        Set ws = Nothing
        Set ws = wb.Worksheets(sheetName)
        If ws Is Nothing Then Exit Do
        idx = idx + 1
        sheetName = "Mau_NoiSuy_2D_" & idx
    Loop
    
    Set ws = wb.Worksheets.Add(After:=wb.ActiveSheet)
    ws.Name = sheetName
    
    ' Tieu de bang
    ws.Range("B2:G2").Merge
    ws.Range("B2").Value = "BANG TRA MA TRAN 2 CHIEU MAU (NOISUYMATRAN)"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 13
    ws.Range("B2").Font.Color = RGB(0, 51, 102)
    
    ws.Range("B3").Value = "He so alpha \ Ty so L/B ->"
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
    
    ' Vien khung bang tra
    ws.Range("B3:G9").Borders.LineStyle = xlContinuous
    
    ' Khu vuc tra cuu
    ws.Range("B12").Value = "TRA CUU NOI SUY MA TRAN:"
    ws.Range("B12").Font.Bold = True
    ws.Range("B12").Font.Color = RGB(150, 50, 0)
    
    ws.Range("B13").Value = "Ty so L/B can tra (Xu):"
    ws.Range("C13").Value = 1.75
    ws.Range("C13").Interior.Color = RGB(255, 255, 200)
    ws.Range("C13").Font.Bold = True
    
    ws.Range("B14").Value = "Chieu sau z/B can tra (Yu):"
    ws.Range("C14").Value = 0.5
    ws.Range("C14").Interior.Color = RGB(255, 255, 200)
    ws.Range("C14").Font.Bold = True
    
    ws.Range("B15").Value = "KET QUA NOI SUY (Alpha):"
    ws.Range("B15").Font.Bold = True
    ws.Range("C15").Formula = "=NoisuyMatran(C$3:G$3, $B$4:$B$9, C$4:G$9, C13, C14)"
    ws.Range("C15").Font.Bold = True
    ws.Range("C15").Font.Size = 12
    ws.Range("C15").Font.Color = RGB(0, 100, 0)
    ws.Range("C15").Interior.Color = RGB(210, 245, 210)
    ws.Range("C15").NumberFormat = "0.0000"
    
    ws.Range("E12").Value = "GIAI THICH CONG THUC:"
    ws.Range("E12").Font.Bold = True
    ws.Range("E13").Value = "=NoisuyMatran(Xk, Yk, GT, Xu, Yu)"
    ws.Range("E13").Font.Bold = True
    ws.Range("E14").Value = "- Xk: Hang moc ngang C3:G3 (Mau xanh duong)"
    ws.Range("E15").Value = "- Yk: Cot moc doc B4:B9 (Mau vang cam)"
    ws.Range("E16").Value = "- GT: Ma tran du lieu C4:G9 (Mau trang)"
    ws.Range("E17").Value = "- Xu, Yu: O gia tri can tra C13, C14"
    
    ws.Columns("B:G").AutoFit
    InsertNoiSuyToActiveWorkbook showMsg:=False
    
    If showMsg Then MsgBox "DA TAO BANG TRA MA TRAN 2D MAU TAI SHEET '" & sheetName & "'!" & vbCrLf & vbCrLf & _
           "Bang da gan san cong thuc =NoisuyMatran(...) hoan chinh de ban tham khao.", vbInformation, "AVTool"
End Sub

Public Sub CreateSampleTableDay(Optional ByVal showMsg As Boolean = True)
    Dim wb As Workbook, ws As Worksheet
    Dim sheetName As String
    Dim idx As Long
    
    On Error Resume Next
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Vui long mo mot file Excel de tao bang mau!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    sheetName = "Mau_NoiSuy_1D"
    idx = 1
    Do While True
        Set ws = Nothing
        Set ws = wb.Worksheets(sheetName)
        If ws Is Nothing Then Exit Do
        idx = idx + 1
        sheetName = "Mau_NoiSuy_1D_" & idx
    Loop
    
    Set ws = wb.Worksheets.Add(After:=wb.ActiveSheet)
    ws.Name = sheetName
    
    ws.Range("B2:D2").Merge
    ws.Range("B2").Value = "BANG TRA DAY SO 1 CHIEU MAU (NOISUYDAY)"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 13
    ws.Range("B2").Font.Color = RGB(0, 51, 102)
    
    ws.Range("B4").Value = "Moc X"
    ws.Range("C4").Value = "Gia Tri Y"
    ws.Range("B4:C4").Interior.Color = RGB(220, 235, 252)
    ws.Range("B4:C4").Font.Bold = True
    ws.Range("B4:C4").HorizontalAlignment = xlCenter
    
    ws.Range("B5:B11").Value = Application.Transpose(Array(10, 20, 30, 40, 50, 75, 100))
    ws.Range("C5:C11").Value = Application.Transpose(Array(1.25, 2.4, 3.65, 4.8, 6.1, 9.25, 12.5))
    ws.Range("B4:C11").Borders.LineStyle = xlContinuous
    ws.Range("B5:C11").HorizontalAlignment = xlRight
    
    ws.Range("E4").Value = "TRA CUU NOI SUY 1 CHIEU:"
    ws.Range("E4").Font.Bold = True
    ws.Range("E5").Value = "Gia tri X can tra (Xu):"
    ws.Range("F5").Value = 35
    ws.Range("F5").Interior.Color = RGB(255, 255, 200)
    ws.Range("F5").Font.Bold = True
    
    ws.Range("E6").Value = "KET QUA NOI SUY (Y):"
    ws.Range("E6").Font.Bold = True
    ws.Range("F6").Formula = "=NoisuyDay($B$5:$B$11, $C$5:$C$11, F5)"
    ws.Range("F6").Font.Bold = True
    ws.Range("F6").Font.Size = 12
    ws.Range("F6").Font.Color = RGB(0, 100, 0)
    ws.Range("F6").Interior.Color = RGB(210, 245, 210)
    ws.Range("F6").NumberFormat = "0.00"
    
    ws.Columns("B:G").AutoFit
    InsertNoiSuyToActiveWorkbook showMsg:=False
    
    If showMsg Then MsgBox "DA TAO BANG TRA DAY SO 1D MAU TAI SHEET '" & sheetName & "'!" & vbCrLf & vbCrLf & _
           "Bang da gan san cong thuc =NoisuyDay(...) hoan chinh de ban tham khao.", vbInformation, "AVTool"
End Sub

Public Sub CreateSampleTableThep(Optional ByVal showMsg As Boolean = True)
    Dim wb As Workbook, ws As Worksheet
    Dim sheetName As String
    Dim idx As Long, r As Long, c As Long
    Dim phiArr As Variant
    
    On Error Resume Next
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Vui long mo mot file Excel de tao bang mau!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    sheetName = "Bang_Tra_Thep"
    idx = 1
    Do While True
        Set ws = Nothing
        Set ws = wb.Worksheets(sheetName)
        If ws Is Nothing Then Exit Do
        idx = idx + 1
        sheetName = "Bang_Tra_Thep_" & idx
    Loop
    
    Set ws = wb.Worksheets.Add(After:=wb.ActiveSheet)
    ws.Name = sheetName
    
    ws.Range("B2:L2").Merge
    ws.Range("B2").Value = "BANG TRA DIEN TICH COT THEP (cm2) THEO DUONG KINH & SO THANH"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 13
    ws.Range("B2").Font.Color = RGB(0, 51, 102)
    
    ws.Range("B4").Value = "Phi (mm)"
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Interior.Color = RGB(220, 235, 252)
    ws.Range("B4").HorizontalAlignment = xlCenter
    
    For c = 1 To 10
        ws.Cells(4, 2 + c).Value = c & " thanh"
        ws.Cells(4, 2 + c).Font.Bold = True
        ws.Cells(4, 2 + c).Interior.Color = RGB(220, 235, 252)
        ws.Cells(4, 2 + c).HorizontalAlignment = xlCenter
    Next c
    
    phiArr = Array(6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 30, 32, 36, 40)
    For r = 0 To UBound(phiArr)
        ws.Cells(5 + r, 2).Value = phiArr(r)
        ws.Cells(5 + r, 2).Font.Bold = True
        ws.Cells(5 + r, 2).Interior.Color = RGB(255, 242, 204)
        ws.Cells(5 + r, 2).HorizontalAlignment = xlCenter
        
        For c = 1 To 10
            ' Dien tich cm2 = c * Dt(phi) * 10000
            ws.Cells(5 + r, 2 + c).Formula = "=" & c & " * Dt($B" & (5 + r) & ") * 10000"
            ws.Cells(5 + r, 2 + c).NumberFormat = "0.00"
        Next c
    Next r
    
    ws.Range("B4:L" & (4 + UBound(phiArr) + 1)).Borders.LineStyle = xlContinuous
    ws.Columns("B:L").AutoFit
    InsertNoiSuyToActiveWorkbook showMsg:=False
    
    If showMsg Then MsgBox "DA TAO BANG TRA DIEN TICH COT THEP TAI SHEET '" & sheetName & "'!" & vbCrLf & vbCrLf & _
           "Bang da tinh toan san tu Phi 6 den Phi 40 cho 1 den 10 thanh bang ham =Dt(...).", vbInformation, "AVTool"
End Sub
