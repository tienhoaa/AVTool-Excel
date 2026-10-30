'========================================================================================
' MODULE: Mod_FontViet_Core
' AVTool - He Thong Chuyen Doi Font Tieng Viet Cho Excel
'========================================================================================

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If


Public gProtectFormula As Boolean
Public gSkipShapes As Boolean
Public gProcessScope As Integer ' 1: Selection, 2: ActiveSheet, 3: AllSheets

Private mtUni As Variant
Private mtAbc As Variant
Private mtVni As Variant
Private bInitTables As Boolean

Public Sub InitSettings()
    gProtectFormula = True  ' Mac dinh LUON bao ve cong thuc
    gSkipShapes = True      ' Mac dinh BO QUA hinh anh va Shapes
    If gProcessScope = 0 Then gProcessScope = 1 ' Mac dinh: Selection
End Sub

Public Sub InitTables()
    If bInitTables Then Exit Sub
    
    mtUni = Array(225, 224, 7843, 227, 7841, 259, 7855, 7857, 7859, 7861, 7863, 226, 7845, 7847, 7849, 7851, 7853, 273, 233, 232, 7867, 7869, 7865, 234, 7871, 7873, 7875, 7877, 7879, 237, 236, 7881, 297, 7883, 243, 242, 7887, 245, 7885, 244, 7889, 7891, 7893, 7895, 7897, 417, 7899, 7901, 7903, 7905, 7907, 250, 249, 7911, 361, 7909, 432, 7913, 7915, 7917, 7919, 7921, 253, 7923, 7927, 7929, 7925, 193, 192, 7842, 195, 7840, 258, 7854, 7856, 7858, 7860, 7862, 194, 7844, 7846, 7848, 7850, 7852, 272, 201, 200, 7866, 7868, 7864, 202, 7870, 7872, 7874, 7876, 7878, 205, 204, 7880, 296, 7882, 211, 210, 7886, 213, 7884, 212, 7888, 7890, 7892, 7894, 7896, 416, 7898, 7900, 7902, 7904, 7906, 218, 217, 7910, 360, 7908, 431, 7912, 7914, 7916, 7918, 7920, 221, 7922, 7926, 7928, 7924)

    mtAbc = Array(184, 181, 182, 183, 185, 168, 190, 187, 188, 189, 198, 169, 202, 199, 200, 201, 203, 174, 208, 204, 206, 207, 209, 170, 213, 210, 211, 212, 214, 221, 215, 216, 220, 222, 227, 223, 225, 226, 228, 171, 232, 229, 230, 231, 233, 172, 237, 234, 235, 236, 238, 243, 239, 241, 242, 244, 173, 248, 245, 246, 247, 249, 253, 250, 251, 252, 254, 184, 181, 182, 183, 185, 161, 190, 187, 188, 189, 198, 162, 202, 199, 200, 201, 203, 167, 208, 204, 206, 207, 209, 163, 213, 210, 211, 212, 214, 221, 215, 216, 220, 222, 227, 223, 225, 226, 228, 164, 232, 229, 230, 231, 233, 165, 237, 234, 235, 236, 238, 243, 239, 241, 242, 244, 166, 248, 245, 246, 247, 244, 253, 250, 251, 252, 254)

    mtVni = Array(97249, 97248, 97251, 97245, 97239, 97234, 97233, 97232, 97250, 97252, 97235, 97226, 97225, 97224, 97229, 97227, 97228, 241, 101249, 101248, 101251, 101245, 101239, 101226, 101225, 101224, 101229, 101227, 101228, 237, 236, 230, 243, 242, 111249, 111248, 111251, 111245, 111239, 111226, 111225, 111224, 111229, 111227, 111228, 244, 244249, 244248, 244251, 244245, 244239, 117249, 117248, 117251, 117245, 117239, 246, 246249, 246248, 246251, 246245, 246239, 121249, 121248, 121251, 121245, 238, 65217, 65216, 65219, 65213, 65207, 65202, 65201, 65200, 65218, 65220, 65203, 65194, 65193, 65192, 65197, 65195, 65196, 209, 69217, 69216, 69219, 69213, 69207, 69194, 69193, 69192, 69197, 69195, 69196, 205, 204, 198, 211, 210, 79217, 79216, 79219, 79213, 79207, 79194, 79193, 79192, 79197, 79195, 79196, 212, 212217, 212216, 212219, 212213, 212207, 85217, 85216, 85219, 85213, 85207, 214, 214217, 214216, 214219, 214213, 214207, 89217, 89216, 89219, 89213, 206)

    bInitTables = True
End Sub

' 1. TCVN3 -> Unicode
Public Function TCVN3ToUnicode(ByVal sInput As String) As String
    If Len(sInput) = 0 Then
        TCVN3ToUnicode = ""
        Exit Function
    End If
    InitTables
    Dim i As Long, j As Long, code As Long, bFound As Boolean
    Dim ch As String, res As String
    res = ""
    For i = 1 To Len(sInput)
        ch = Mid$(sInput, i, 1)
        code = AscW(ch)
        bFound = False
        For j = 0 To UBound(mtAbc)
            If mtAbc(j) = code Then
                res = res & ChrW$(mtUni(j))
                bFound = True
                Exit For
            End If
        Next j
        If Not bFound Then res = res & ch
    Next i
    TCVN3ToUnicode = res
End Function

' 2. Unicode -> TCVN3
Public Function UnicodeToTCVN3(ByVal sInput As String) As String
    If Len(sInput) = 0 Then
        UnicodeToTCVN3 = ""
        Exit Function
    End If
    InitTables
    Dim i As Long, j As Long, code As Long, bFound As Boolean
    Dim ch As String, res As String
    res = ""
    For i = 1 To Len(sInput)
        ch = Mid$(sInput, i, 1)
        code = AscW(ch)
        bFound = False
        For j = 0 To UBound(mtUni)
            If mtUni(j) = code Then
                res = res & ChrW$(mtAbc(j))
                bFound = True
                Exit For
            End If
        Next j
        If Not bFound Then res = res & ch
    Next i
    UnicodeToTCVN3 = res
End Function

' 3. VNI -> Unicode
Public Function VNIToUnicode(ByVal sInput As String) As String
    If Len(sInput) = 0 Then
        VNIToUnicode = ""
        Exit Function
    End If
    InitTables
    Dim i As Long, j As Long, code1 As Long, code2 As Long
    Dim ch1 As String, ch2 As String, res As String
    Dim val2Char As Double, bFound As Boolean, isDouble As Boolean
    
    res = ""
    i = 1
    Do While i <= Len(sInput)
        ch1 = Mid$(sInput, i, 1)
        code1 = AscW(ch1)
        bFound = False
        isDouble = False
        
        If i < Len(sInput) Then
            ch2 = Mid$(sInput, i + 1, 1)
            code2 = AscW(ch2)
            If (code2 >= 224 And code2 <= 254) Or (code2 >= 192 And code2 <= 220) Then
                val2Char = CDbl(code1) * 1000# + CDbl(code2)
                For j = 0 To UBound(mtVni)
                    If CDbl(mtVni(j)) = val2Char Then
                        res = res & ChrW$(mtUni(j))
                        bFound = True
                        isDouble = True
                        Exit For
                    End If
                Next j
            End If
        End If
        
        If Not bFound Then
            For j = 0 To UBound(mtVni)
                If CDbl(mtVni(j)) = CDbl(code1) Then
                    res = res & ChrW$(mtUni(j))
                    bFound = True
                    Exit For
                End If
            Next j
        End If
        
        If bFound Then
            If isDouble Then i = i + 2 Else i = i + 1
        Else
            res = res & ch1
            i = i + 1
        End If
    Loop
    VNIToUnicode = res
End Function

' 4. Unicode -> VNI
Public Function UnicodeToVNI(ByVal sInput As String) As String
    If Len(sInput) = 0 Then
        UnicodeToVNI = ""
        Exit Function
    End If
    InitTables
    Dim i As Long, j As Long, code As Long, bFound As Boolean
    Dim ch As String, res As String, vniCode As Double, c1 As Long, c2 As Long
    res = ""
    For i = 1 To Len(sInput)
        ch = Mid$(sInput, i, 1)
        code = AscW(ch)
        bFound = False
        For j = 0 To UBound(mtUni)
            If mtUni(j) = code Then
                vniCode = CDbl(mtVni(j))
                If vniCode > 65535 Then
                    c1 = Fix(vniCode / 1000#)
                    c2 = CLng(vniCode - (c1 * 1000#))
                    res = res & ChrW$(c1) & ChrW$(c2)
                Else
                    res = res & ChrW$(CLng(vniCode))
                End If
                bFound = True
                Exit For
            End If
        Next j
        If Not bFound Then res = res & ch
    Next i
    UnicodeToVNI = res
End Function

' 5. Tu dong nhan dien TCVN3 hay VNI va chuyen sang Unicode
Public Function AutoToUnicode(ByVal sInput As String, ByVal fontName As String) As String
    Dim uFont As String
    uFont = UCase$(fontName)
    
    If Left$(uFont, 3) = ".VN" Or uFont = "MS SANS SERIF" Or uFont = "TCVN3 SAMPLE FONT" Then
        AutoToUnicode = TCVN3ToUnicode(sInput)
        Exit Function
    ElseIf Left$(uFont, 3) = "VNI" Then
        AutoToUnicode = VNIToUnicode(sInput)
        Exit Function
    End If
    
    Dim i As Long, code As Long
    Dim countAbc As Long, countVni As Long
    For i = 1 To Len(sInput)
        code = AscW(Mid$(sInput, i, 1))
        If (code >= 161 And code <= 174) Or (code >= 181 And code <= 254) Then
            countAbc = countAbc + 1
        End If
        If (code >= 224 And code <= 254) Or (code >= 192 And code <= 220) Then
            countVni = countVni + 1
        End If
    Next i
    
    If countAbc > countVni And countAbc > 0 Then
        AutoToUnicode = TCVN3ToUnicode(sInput)
    ElseIf countVni > 0 Then
        AutoToUnicode = VNIToUnicode(sInput)
    Else
        AutoToUnicode = sInput
    End If
End Function

' 6. Doi hoa, thuong, tieu de Unicode
Public Function ChangeTextCase(ByVal sInput As String, ByVal caseType As Integer) As String
    If Len(sInput) = 0 Then
        ChangeTextCase = ""
        Exit Function
    End If
    Select Case caseType
        Case 1: ChangeTextCase = StrConv(sInput, vbUpperCase)
        Case 2: ChangeTextCase = StrConv(sInput, vbLowerCase)
        Case 3: ChangeTextCase = StrConv(sInput, vbProperCase)
    End Select
End Function

'========================================================================================
' CORE ENGINE: THUC THI CHUYEN DOI AN TOAN TREN EXCEL
'========================================================================================

Public Sub ConvertRange(ByVal targetRange As Range, ByVal convMode As Integer, Optional ByVal caseType As Integer = 0)
    On Error Resume Next
    If targetRange Is Nothing Then Exit Sub
    
    Dim textCells As Range
    ' DIEM COT LOI: Chi lay o la Constant Text, BO QUA HOAN TOAN CAC O CONG THUC
    If gProtectFormula Then
        Set textCells = targetRange.SpecialCells(xlCellTypeConstants, xlTextValues)
        If textCells Is Nothing Then Exit Sub
    Else
        Set textCells = targetRange
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim area As Range, cell As Range
    Dim val As String, newVal As String, fName As String
    
    For Each area In textCells.Areas
        For Each cell In area
            ' Bao ve tuyet doi: Bo qua neu o co cong thuc
            If gProtectFormula And cell.HasFormula Then
                ' Khong lam gi ca
            Else
                val = cell.Value2
                If Len(val) > 0 Then
                    fName = cell.Font.Name
                    Select Case convMode
                        Case 1: newVal = TCVN3ToUnicode(val): cell.Font.Name = "Times New Roman"
                        Case 2: newVal = VNIToUnicode(val): cell.Font.Name = "Times New Roman"
                        Case 3: newVal = AutoToUnicode(val, fName): cell.Font.Name = "Times New Roman"
                        Case 4: newVal = UnicodeToTCVN3(val): cell.Font.Name = ".VnTime"
                        Case 5: newVal = UnicodeToVNI(val): cell.Font.Name = "VNI-Times"
                        Case Else: newVal = val
                    End Select
                    
                    If caseType > 0 Then
                        newVal = ChangeTextCase(newVal, caseType)
                    End If
                    
                    If newVal <> val Then
                        cell.Value = newVal
                    End If
                End If
            End If
        Next cell
    Next area
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub ConvertShapes(ByVal ws As Worksheet, ByVal convMode As Integer, Optional ByVal caseType As Integer = 0)
    If gSkipShapes Then Exit Sub
    
    On Error Resume Next
    Dim shp As Shape
    Dim txt As String, newTxt As String, fName As String
    
    For Each shp In ws.Shapes
        ' BẢO VỆ HÌNH ẢNH: Bỏ qua 13=Picture, 11=LinkedPicture, 12=OLE, 3=Chart
        If shp.Type <> 13 And shp.Type <> 11 And shp.Type <> 12 And shp.Type <> 3 Then
            If shp.HasTextFrame Then
                If shp.TextFrame2.HasText Then
                    txt = shp.TextFrame2.TextRange.Text
                    If Len(txt) > 0 Then
                        fName = shp.TextFrame2.TextRange.Font.Name
                        Select Case convMode
                            Case 1: newTxt = TCVN3ToUnicode(txt): shp.TextFrame2.TextRange.Font.Name = "Times New Roman"
                            Case 2: newTxt = VNIToUnicode(txt): shp.TextFrame2.TextRange.Font.Name = "Times New Roman"
                            Case 3: newTxt = AutoToUnicode(txt, fName): shp.TextFrame2.TextRange.Font.Name = "Times New Roman"
                            Case 4: newTxt = UnicodeToTCVN3(txt): shp.TextFrame2.TextRange.Font.Name = ".VnTime"
                            Case 5: newTxt = UnicodeToVNI(txt): shp.TextFrame2.TextRange.Font.Name = "VNI-Times"
                            Case Else: newTxt = txt
                        End Select
                        If caseType > 0 Then
                            newTxt = ChangeTextCase(newTxt, caseType)
                        End If
                        If newTxt <> txt Then shp.TextFrame2.TextRange.Text = newTxt
                    End If
                End If
            End If
        End If
    Next shp
End Sub

Public Sub ExecuteConversion(ByVal convMode As Integer, Optional ByVal bSilent As Boolean = False, Optional ByVal caseType As Integer = 0)
    InitSettings
    
    Dim msgScope As String
    
    Select Case gProcessScope
        Case 1 ' Selection
            If TypeName(Selection) <> "Range" Then
                If Not bSilent Then MsgBox "Vui long chon mot vung o tren bang tinh!", vbExclamation, "AVTool"
                Exit Sub
            End If
            ConvertRange Selection, convMode, caseType
            msgScope = "vung dang chon"
            
        Case 2 ' ActiveSheet
            ConvertRange ActiveSheet.UsedRange, convMode, caseType
            ConvertShapes ActiveSheet, convMode, caseType
            msgScope = "Sheet hien tai (" & ActiveSheet.Name & ")"
            
        Case 3 ' All Sheets
            Dim ws As Worksheet
            For Each ws In ActiveWorkbook.Worksheets
                ConvertRange ws.UsedRange, convMode, caseType
                ConvertShapes ws, convMode, caseType
            Next ws
            msgScope = "toan bo cac Sheet trong Workbook"
    End Select
    
    Dim modeDesc As String
    Select Case convMode
        Case 1: modeDesc = "TCVN3 (.VnTime) sang Unicode"
        Case 2: modeDesc = "VNI-Windows sang Unicode"
        Case 3: modeDesc = "Tu dong nhan dien sang Unicode"
        Case 4: modeDesc = "Unicode sang TCVN3"
        Case 5: modeDesc = "Unicode sang VNI"
        Case Else: modeDesc = "Doi kieu chu"
    End Select
    
    If Not bSilent Then
        MsgBox "Da hoan thanh chuyen doi " & modeDesc & " cho " & msgScope & "!" & vbCrLf & _
               "- Cong thuc: Da bao ve an toan 100%" & vbCrLf & _
               "- Hinh anh/Shape: " & IIf(gSkipShapes, "Da bo qua an toan", "Da xu ly Text, giu nguyen hinh anh"), _
               vbInformation, "AVTool - Thanh Cong"
    End If
End Sub

Public Sub Run_AutoToUnicode(Optional ByVal bSilent As Boolean = False)
    ExecuteConversion 3, bSilent
End Sub

Public Sub Run_TCVN3ToUnicode(Optional ByVal bSilent As Boolean = False)
    ExecuteConversion 1, bSilent
End Sub

Public Sub Run_VNIToUnicode(Optional ByVal bSilent As Boolean = False)
    ExecuteConversion 2, bSilent
End Sub

Public Sub Run_UnicodeToTCVN3(Optional ByVal bSilent As Boolean = False)
    ExecuteConversion 4, bSilent
End Sub

Public Sub Run_UnicodeToVNI(Optional ByVal bSilent As Boolean = False)
    ExecuteConversion 5, bSilent
End Sub

Public Sub Run_UpperCase()
    ProcessCase 1
End Sub

Public Sub Run_LowerCase()
    ProcessCase 2
End Sub

Public Sub Run_TitleCase()
    ProcessCase 3
End Sub

Private Sub ProcessCase(ByVal caseType As Integer)
    InitSettings
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range, textCells As Range
    On Error Resume Next
    Set textCells = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    If textCells Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    For Each cell In textCells
        If Not (gProtectFormula And cell.HasFormula) Then
            cell.Value = ChangeTextCase(cell.Value2, caseType)
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub ShowAVToolForm()
    On Error Resume Next
    frmAVTool.Show vbModeless
    If Err.Number <> 0 Then
        frmAVTool.Show
    End If
End Sub

Public Sub ShowHelpDialog()
    MsgBox "AVTOOL - HE THONG CHUYEN DOI FONT TIENG VIET PRO" & vbCrLf & vbCrLf & _
           "1. An toan cong thuc:" & vbCrLf & _
           "   - Mac dinh LUON bao ve cong thuc. Cac o chua cong thuc (=SUM, =VLOOKUP...) se khong bao gio bi sua doi." & vbCrLf & vbCrLf & _
           "2. An toan hinh anh:" & vbCrLf & _
           "   - Nut 'Bo qua Hinh anh / Shape' loai tru 100% hinh anh, logo, chu ky scan de tranh loi va xo lech." & vbCrLf & vbCrLf & _
           "3. Ho tro day du tren Excel 2010 - 2021, Office 365.", _
           vbInformation, "AVTool - Tro Giup"
End Sub

' RIBBON CALLBACKS
Public Sub Ribbon_OnAction(control As IRibbonControl)
    Select Case control.Id
        ' Form
        Case "btnOpenForm": ShowAVToolForm
        Case "btnInsertDValue": InsertDValueToActiveWorkbook
        Case "btnBatchPrint": ShowBatchPrintForm
        Case "btnCheckUpdate": CheckForUpdates False
        
        ' Scope
        Case "btnScopeSelection": gProcessScope = 1: MsgBox "Da dat pham vi: Vung dang chon (Selection)", vbInformation, "AVTool"
        Case "btnScopeSheet": gProcessScope = 2: MsgBox "Da dat pham vi: Sheet hien tai", vbInformation, "AVTool"
        Case "btnScopeWorkbook": gProcessScope = 3: MsgBox "Da dat pham vi: Toan bo Workbook", vbInformation, "AVTool"
        
        ' Convert
        Case "btnAutoToUni": Run_AutoToUnicode
        Case "btnAbcToUni": Run_TCVN3ToUnicode
        Case "btnVniToUni": Run_VNIToUnicode
        Case "btnUniToAbc": Run_UnicodeToTCVN3
        Case "btnUniToVni": Run_UnicodeToVNI
        
        ' Safety Toggles
        Case "btnToggleFormula":
            gProtectFormula = Not gProtectFormula
            If gProtectFormula Then
                MsgBox "Bao ve cong thuc: DANG BAT" & vbCrLf & "- Tat ca cac o chua cong thuc duoc giu nguyen 100%.", vbInformation, "AVTool"
            Else
                MsgBox "Bao ve cong thuc: DANG TAT" & vbCrLf & "- Canh bao: Cac o co the bi ghi de gia tri!", vbExclamation, "AVTool"
            End If
        Case "btnToggleShapes":
            gSkipShapes = Not gSkipShapes
            If gSkipShapes Then
                MsgBox "Bo qua Hinh anh & Shapes: DANG BAT" & vbCrLf & "- Tuyet doi khong cham vao bat ky hinh anh/shape nao.", vbInformation, "AVTool"
            Else
                MsgBox "Bo qua Hinh anh & Shapes: DANG TAT" & vbCrLf & "- Se doi Text trong Shapes (van tu dong bao ve Picture).", vbInformation, "AVTool"
            End If
            
        ' Case
        Case "btnUpper": Run_UpperCase
        Case "btnLower": Run_LowerCase
        Case "btnTitle": Run_TitleCase
        
        ' Help
        Case "btnHelp": ShowHelpDialog
    End Select
End Sub

Private Sub Auto_Open()
    InitSettings
End Sub



' INSERT DVALUE MODULE INTO ACTIVE WORKBOOK

Public Sub InsertDValueToActiveWorkbook()
    Dim targetWb As Workbook
    Dim comp As Object
    Dim existingComp As Object
    Dim vbProj As Object
    Dim dvalCode As String
    Dim answer As VbMsgBoxResult
    Dim srcComp As Object
    Dim lineCount As Long
    Dim initName As String
    Dim vFile As Variant
    Dim codeText As String
    
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    
    Set targetWb = ActiveWorkbook
    If targetWb Is Nothing Then
        MsgBox "Chua co file Excel nao dang mo!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    If targetWb Is ThisWorkbook Then
        MsgBox "Khong the chen vao chinh file Add-in!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    ' Check VBProject access
    Err.Clear
    Set vbProj = targetWb.VBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        MsgBox "Khong the chen Module vao file vi Excel chua mo khoa quyen VBA." & vbCrLf & vbCrLf & _
               "Cach bat rat nhanh (chi can lam 1 lan):" & vbCrLf & _
               "1. Vao File > Options > Trust Center > Trust Center Settings..." & vbCrLf & _
               "2. Chon Macro Settings > Tich chon 'Trust access to the VBA project object model' > OK.", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    ' Read source code from ThisWorkbook Mod_DValue
    Set srcComp = ThisWorkbook.VBProject.VBComponents("Mod_DValue")
    If srcComp Is Nothing Then
        MsgBox "Loi: Khong tim thay ma nguon Mod_DValue trong Add-in!", vbCritical, "AVTool"
        Exit Sub
    End If
    
    lineCount = srcComp.CodeModule.CountOfLines
    dvalCode = srcComp.CodeModule.Lines(1, lineCount)
    
    ' Smart scan: check if DValue function ALREADY EXISTS in ANY module of target workbook
    Set existingComp = Nothing
    For Each comp In vbProj.VBComponents
        If comp.Type = 1 Then ' vbext_ct_StdModule
            lineCount = comp.CodeModule.CountOfLines
            If lineCount > 0 Then
                codeText = comp.CodeModule.Lines(1, lineCount)
                If InStr(1, codeText, "Function DValue", vbTextCompare) > 0 Then
                    Set existingComp = comp
                    Exit For
                End If
            End If
        End If
    Next comp
    
    If Not existingComp Is Nothing Then
        answer = MsgBox("Trong file '" & targetWb.Name & "' da co san ham DValue tai Module '" & existingComp.Name & "'!" & vbCrLf & vbCrLf & _
                        "Ban co muon CAP NHAT lai ma nguon DValue ban moi nhat vao Module nay khong?" & vbCrLf & _
                        "- Chon YES de cap nhat an toan (ghi de code moi, khong bi trung ten ham)." & vbCrLf & _
                        "- Chon NO de giu nguyen code cu va chuyen sang Luu file.", vbQuestion + vbYesNo, "AVTool")
        If answer = vbYes Then
            ' Overwrite cleanly without deleting component to avoid reset/break
            existingComp.CodeModule.DeleteLines 1, existingComp.CodeModule.CountOfLines
            existingComp.CodeModule.AddFromString dvalCode
        Else
            GoTo PromptSaveAs
        End If
    Else
        ' Create new standard module Mod_DValue
        Set comp = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
        comp.Name = "Mod_DValue"
        comp.CodeModule.AddFromString dvalCode
    End If
    
    ' Close VBA editor if open
    On Error Resume Next
    If Not Application.VBE Is Nothing Then
        Application.VBE.MainWindow.Visible = False
    End If
    
    ' Recalculate
    Application.Calculate
    
PromptSaveAs:
    ' Save As dialog
    initName = targetWb.Name
    If LCase$(Right$(initName, 5)) = ".xlsx" Then
        initName = Left$(initName, Len(initName) - 5) & ".xls"
    ElseIf LCase$(Right$(initName, 5)) = ".xlsm" Then
        ' Keep .xlsm
    ElseIf LCase$(Right$(initName, 4)) = ".xls" Then
        ' Keep .xls
    ElseIf InStr(initName, ".") = 0 Then
        initName = initName & ".xls"
    End If
    
    vFile = Application.GetSaveAsFilename( _
        InitialFileName:=initName, _
        FileFilter:="Excel 97-2003 Workbook (*.xls), *.xls, Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
        FilterIndex:=1, _
        Title:="AVTool - Luu file ho tro Macro DValue (Nen luu .xls hoac .xlsm)")
        
    If vFile <> False Then
        If LCase$(Right$(vFile, 4)) = ".xls" Then
            targetWb.SaveAs Filename:=vFile, FileFormat:=56 ' xlExcel8 (.xls)
        ElseIf LCase$(Right$(vFile, 5)) = ".xlsm" Then
            targetWb.SaveAs Filename:=vFile, FileFormat:=52 ' xlOpenXMLWorkbookMacroEnabled (.xlsm)
        Else
            targetWb.SaveAs Filename:=vFile
        End If
        MsgBox "DA CHEN VA LUU FILE THANH CONG!" & vbCrLf & vbCrLf & _
               "File: " & targetWb.Name & vbCrLf & _
               "Ham DValue da san sang va duoc luu giu vinh vien trong file nay.", vbInformation, "AVTool"
    Else
        MsgBox "DA CHEN HAM DVALUE VAO FILE THANH CONG!" & vbCrLf & vbCrLf & _
               "Luu y: Ban chua luu file. Hay nho luu file dang .xls hoac .xlsm de giu lai ham DValue nhe!", vbInformation, "AVTool"
    End If
End Sub


' =========================================================
' AUTO-UPDATE ONLINE FROM GITHUB REPOSITORY
' =========================================================
Public Sub Auto_Open()
    On Error Resume Next
    ' Tu dong kiem tra cap nhat sau 3 giay khi mo Excel
    Application.OnTime Now + TimeValue("00:00:03"), "CheckForUpdatesSilent"
End Sub

Public Sub CheckForUpdatesSilent()
    CheckForUpdates True
End Sub

Public Sub CheckForUpdates(Optional ByVal bSilent As Boolean = False)
    Dim http As Object
    Dim jsonStr As String
    Dim remoteVer As String
    Dim changeLog As String
    Dim downUrl As String
    Dim currentVer As String
    Dim ans As VbMsgBoxResult
    
    currentVer = "1.0.2"
    
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", "https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/version.json?t=" & Timer, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.setTimeouts 3000, 3000, 3000, 3000
    http.Send
    
    If http.Status <> 200 Then
        If Not bSilent Then
            MsgBox "Khong the ket noi toi may chu cap nhat GitHub." & vbCrLf & _
                   "Vui long kiem tra lai ket noi Internet hoac truy cap truc tiep:" & vbCrLf & _
                   "https://github.com/tienhoaa/AVTool-Excel", vbExclamation, "AVTool Update"
        End If
        Exit Sub
    End If
    
    jsonStr = http.responseText
    remoteVer = ExtractJsonVal(jsonStr, "latest_version")
    If Len(remoteVer) = 0 Then remoteVer = ExtractJsonVal(jsonStr, "version")
    changeLog = ExtractJsonVal(jsonStr, "changelog")
    downUrl = ExtractJsonVal(jsonStr, "download_url")
    If Len(downUrl) = 0 Then downUrl = "https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/TN_FontViet_Pro_2026.xlam"
    
    If Len(remoteVer) = 0 Then Exit Sub
    
    If CompareVersions(remoteVer, currentVer) > 0 Then
        ans = MsgBox("PHAT HIEN PHIEN BAN MOI CHO AVTOOL PRO!" & vbCrLf & _
                     "--------------------------------------------------" & vbCrLf & _
                     "- Phien ban ban dang dung: v" & currentVer & vbCrLf & _
                     "- Phien ban moi nhat: v" & remoteVer & vbCrLf & vbCrLf & _
                     "Noi dung tinh nang moi:" & vbCrLf & changeLog & vbCrLf & _
                     "--------------------------------------------------" & vbCrLf & _
                     "Ban co muon tu dong cap nhat ngay bay gio khong?", vbQuestion + vbYesNo, "AVTool Auto-Update")
        If ans = vbYes Then
            DoAutoUpdateInPlace downUrl, remoteVer
        End If
    Else
        If Not bSilent Then
            MsgBox "Ban dang su dung phien ban moi nhat (v" & currentVer & ")!" & vbCrLf & _
                   "He thong luon san sang va da toi uu nhat.", vbInformation, "AVTool Update"
        End If
    End If
End Sub

Private Sub DoAutoUpdateInPlace(ByVal downloadUrl As String, ByVal newVer As String)
    Dim tempXlam As String, batPath As String, targetPath As String
    Dim res As Long, fNum As Integer
    
    targetPath = ThisWorkbook.FullName
    tempXlam = Environ$("TEMP") & "\AVTool_Update.xlam"
    batPath = Environ$("TEMP") & "\AVTool_ApplyUpdate.bat"
    
    On Error Resume Next
    If Dir$(tempXlam) <> "" Then Kill tempXlam
    If Dir$(batPath) <> "" Then Kill batPath
    
    Application.StatusBar = "Dang tai ban cap nhat moi tu GitHub, vui long doi..."
    
    ' 1. Tai file truc tiep bang Windows API URLDownloadToFile (100% an toan)
    res = URLDownloadToFile(0, downloadUrl, tempXlam, 0, 0)
    
    If res <> 0 Or Dir$(tempXlam) = "" Then
        MsgBox "Khong the tai ban cap nhat moi tu GitHub. Vui long kiem tra ket noi Internet!", vbCritical, "AVTool Update"
        Application.StatusBar = False
        Exit Sub
    End If
    
    If FileLen(tempXlam) < 20000 Then
        MsgBox "Tep tin tai ve bi loi hoac khong hoan tat!", vbCritical, "AVTool Update"
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' 2. Tao script copy de file sau khi Excel dong (chi co timeout & copy & start excel)
    fNum = FreeFile
    Open batPath For Output As #fNum
    Print #fNum, "@echo off"
    Print #fNum, "timeout /t 1 /nobreak >nul"
    Print #fNum, "copy /y """ & tempXlam & """ """ & targetPath & """ >nul"
    Print #fNum, "del /f /q """ & tempXlam & """ >nul"
    Print #fNum, "start excel.exe"
    Print #fNum, "(goto) 2>nul & del ""%~f0"""
    Close #fNum
    
    ' 3. Thuc thi cap nhat va thoat Excel de ghi de
    Shell Chr(34) & batPath & Chr(34), vbHide
    ThisWorkbook.Saved = True
    Application.Quit
End Sub

Private Function ExtractJsonVal(ByVal json As String, ByVal key As String) As String
    Dim p1 As Long, p2 As Long, searchKey As String
    searchKey = """" & key & """"
    p1 = InStr(1, json, searchKey, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = InStr(p1 + Len(searchKey), json, ":")
    If p1 = 0 Then Exit Function
    p1 = InStr(p1 + 1, json, """")
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, json, """")
    If p2 = 0 Then Exit Function
    ExtractJsonVal = Mid$(json, p1 + 1, p2 - p1 - 1)
End Function

Private Function CompareVersions(ByVal v1 As String, ByVal v2 As String) As Long
    Dim p1() As String, p2() As String
    Dim i As Long, n1 As Long, n2 As Long
    p1 = Split(v1, ".")
    p2 = Split(v2, ".")
    For i = 0 To UBound(p1)
        If i > UBound(p2) Then CompareVersions = 1: Exit Function
        n1 = Val(p1(i)): n2 = Val(p2(i))
        If n1 > n2 Then CompareVersions = 1: Exit Function
        If n1 < n2 Then CompareVersions = -1: Exit Function
    Next i
    If UBound(p2) > UBound(p1) Then CompareVersions = -1 Else CompareVersions = 0
End Function
