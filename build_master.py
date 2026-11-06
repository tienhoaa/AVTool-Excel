import os, sys, zipfile, subprocess, shutil, time

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
bas_core = os.path.join(workspace, "Mod_FontViet_Core.bas")
bas_dval = os.path.join(workspace, "Mod_DValue.bas")
bas_print = os.path.join(workspace, "Mod_BatchPrint.bas")
bas_noisuy = os.path.join(workspace, "Mod_NoiSuy.bas")
temp_xlam = os.path.join(workspace, "temp_addon.xlam")
final_xlam = os.path.join(workspace, "TN_FontViet_Pro_2026.xlam")

subprocess.run(["powershell", "-Command", "Stop-Process -Name 'EXCEL' -Force -ErrorAction SilentlyContinue"])
time.sleep(1)

def esc(text):
    out = []
    for ch in text:
        if ord(ch) > 127:
            out.append(f"\\u{ord(ch):04X}")
        else:
            out.append(ch)
    return "".join(out)

# Form 1: frmAVTool Captions
elements_avtool = [
    ("Me.Caption", "AVTool - Chuyen Doi Font Tieng Viet Pro"),
    ("lblHeader.Caption", "AVTOOL - CHUYỂN ĐỔI FONT TIẾNG VIỆT"),
    ("lblSub.Caption", "Tự động nhận diện TCVN3, VNI → Unicode | Bảo toàn 100% công thức"),
    ("fraSource.Caption", " Bảng mã nguồn "),
    ("optSrcAuto.Caption", "Tự động nhận diện (Khuyên dùng)"),
    ("optSrcAbc.Caption", "TCVN3 (Font .VnTime, .VnArial...)"),
    ("optSrcVni.Caption", "VNI-Windows (Font VNI-Times...)"),
    ("optSrcUni.Caption", "Unicode"),
    ("fraTarget.Caption", " Bảng mã đích "),
    ("optDstUni.Caption", "Unicode (Chuẩn quốc tế)"),
    ("optDstAbc.Caption", "TCVN3 (.VnTime)"),
    ("optDstVni.Caption", "VNI-Windows"),
    ("fraScope.Caption", " Phạm vi áp dụng "),
    ("optScopeSelection.Caption", "Vùng đang chọn"),
    ("optScopeSheet.Caption", "Sheet hiện tại"),
    ("optScopeWorkbook.Caption", "Toàn bộ File"),
    ("fraSafety.Caption", " Tùy chọn an toàn & Bảo vệ "),
    ("chkFormula.Caption", "Bảo vệ công thức (Tuyệt đối không mất công thức)"),
    ("chkShapes.Caption", "Bỏ qua hình ảnh & đồ họa (Giữ nguyên ảnh/Shapes)"),
    ("fraCase.Caption", " Kiểu chữ (Tùy chọn) "),
    ("optCaseNone.Caption", "Giữ nguyên"),
    ("optCaseUpper.Caption", "IN HOA TOÀN BỘ"),
    ("optCaseLower.Caption", "in thường"),
    ("optCaseTitle.Caption", "Viết Hoa Đầu Từ"),
    ("btnExecute.Caption", "THỰC HIỆN CHUYỂN ĐỔI"),
    ("btnClose.Caption", "Đóng")
]

lines_avtool = []
for var, val in elements_avtool:
    if var == "Me.Caption":
        lines_avtool.append(f'    {var} = "{val}"')
    else:
        lines_avtool.append(f'    {var} = UniText("{esc(val)}")')
captions_avtool = "\n".join(lines_avtool)

code_frmAVTool = f'''#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Function UniText(ByVal s As String) As String
    Dim pos As Long, hexStr As String, res As String
    res = s
    pos = InStr(res, "\\u")
    Do While pos > 0
        hexStr = Mid(res, pos + 2, 4)
        res = Left(res, pos - 1) & ChrW(CLng("&H" & hexStr)) & Mid(res, pos + 6)
        pos = InStr(res, "\\u")
    Loop
    UniText = res
End Function

Private Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "AVTool") As Long
    MsgBoxUni = MessageBoxW(0, StrPtr(UniText(prompt)), StrPtr(UniText(title)), buttons)
End Function

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    optSrcAuto.Value = True
    optDstUni.Value = True
    optScopeSelection.Value = True
    chkFormula.Value = True
    chkShapes.Value = True
    optCaseNone.Value = True
    
    Dim c As Control
    For Each c In Me.Controls
        c.Font.Name = "Segoe UI"
        c.Font.Size = 9
    Next c
    lblHeader.Font.Size = 11
    lblHeader.Font.Bold = True
    lblSub.Font.Size = 8.5
    btnExecute.Font.Bold = True
    btnExecute.Font.Size = 9.5
    btnClose.Font.Size = 9.5
    
{captions_avtool}
End Sub

Private Sub btnExecute_Click()
    Dim srcMode As Integer
    Dim dstMode As Integer
    Dim caseType As Integer
    
    If optScopeSelection.Value Then
        gProcessScope = 1
    ElseIf optScopeSheet.Value Then
        gProcessScope = 2
    ElseIf optScopeWorkbook.Value Then
        gProcessScope = 3
    End If
    
    gProtectFormula = chkFormula.Value
    gSkipShapes = chkShapes.Value
    
    ' Bảng mã nguồn: 0: Tự động, 1: Unicode, 2: TCVN3, 3: VNI
    If optSrcAuto.Value Then
        srcMode = 0
    ElseIf optSrcUni.Value Then
        srcMode = 1
    ElseIf optSrcAbc.Value Then
        srcMode = 2
    ElseIf optSrcVni.Value Then
        srcMode = 3
    Else
        srcMode = 0
    End If
    
    ' Bảng mã đích: 1: Unicode, 2: TCVN3, 3: VNI
    If optDstUni.Value Then
        dstMode = 1
    ElseIf optDstAbc.Value Then
        dstMode = 2
    ElseIf optDstVni.Value Then
        dstMode = 3
    Else
        dstMode = 1
    End If
    
    ' Kiểu chữ
    If optCaseUpper.Value Then
        caseType = 1
    ElseIf optCaseLower.Value Then
        caseType = 2
    ElseIf optCaseTitle.Value Then
        caseType = 3
    Else
        caseType = 0
    End If
    
    If srcMode <> 0 And srcMode = dstMode And caseType = 0 Then
        MsgBoxUni "{esc('Bảng mã nguồn và bảng mã đích đang giống nhau! Vui lòng chọn bảng mã khác hoặc kiểu chữ cần đổi.')}", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    Me.Hide
    ExecuteConversionFull srcMode, dstMode, False, caseType
    Me.Show vbModeless
End Sub
'''

# Form 2: Clean frmBatchPrint Captions & Code
cbo_item0 = esc("1. Theo từng Sheet (Mặc định)")
cbo_item1 = esc("2. Đồng loạt khổ A4")
cbo_item2 = esc("3. Đồng loạt khổ A3")

elements_batchprint = [
    ("Me.Caption", "AVTool - In Nhieu Sheet (Fit All Columns & Paper Size)"),
    ("lblHeader.Caption", "AVTOOL - IN NHIỀU SHEET (BATCH PRINT)"),
    ("lblSub.Caption", "Chọn sheet in, co vừa chiều ngang trang (Fit All Columns) & chọn máy in"),
    ("fraSource.Caption", " Sheet trong Workbook "),
    ("fraTarget.Caption", " Danh sách Sheet cần in "),
    ("chkAllVisible.Caption", "Chỉ hiển thị các Sheet không bị ẩn"),
    ("lblTabColor.Caption", "Chọn nhanh theo màu Tab:"),
    ("btnApplyColor.Caption", "Chuyển sang in >>"),
    ("fraOptions.Caption", " Tùy chọn in ấn & Máy in "),
    ("lblPaper.Caption", "Khổ giấy:"),
    ("chkFitAllColumns.Caption", "Co vừa chiều ngang 1 trang (Fit All Columns on One Page)"),
    ("lblPrinterTitle.Caption", "Máy in:"),
    ("btnChangePrinter.Caption", "Đổi máy in..."),
    ("btnPreview.Caption", "Xem trước (Print Preview)"),
    ("btnPrint.Caption", "In ngay (Print Out)"),
    ("btnClose.Caption", "Đóng"),
    ("btnAddOne.Caption", ">"),
    ("btnAddAll.Caption", ">>"),
    ("btnRemoveOne.Caption", "<"),
    ("btnRemoveAll.Caption", "<<"),
    ("btnMoveUp.Caption", "Lên"),
    ("btnMoveDown.Caption", "Xuống")
]

lines_batchprint = []
for var, val in elements_batchprint:
    if var == "Me.Caption":
        lines_batchprint.append(f'    {var} = "{val}"')
    else:
        lines_batchprint.append(f'    {var} = UniText("{esc(val)}")')
captions_batchprint = "\n".join(lines_batchprint)

code_frmBatchPrint = f'''#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Function UniText(ByVal s As String) As String
    Dim pos As Long, hexStr As String, res As String
    res = s
    pos = InStr(res, "\\u")
    Do While pos > 0
        hexStr = Mid(res, pos + 2, 4)
        res = Left(res, pos - 1) & ChrW(CLng("&H" & hexStr)) & Mid(res, pos + 6)
        pos = InStr(res, "\\u")
    Loop
    UniText = res
End Function

Private Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    If Len(title) = 0 Then title = UniText("{esc('AVTool - In Nhiều Sheet')}")
    MsgBoxUni = MessageBoxW(0, StrPtr(UniText(prompt)), StrPtr(UniText(title)), buttons)
End Function

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim c As Control
    For Each c In Me.Controls
        c.Font.Name = "Segoe UI"
        c.Font.Size = 9
    Next c
    lblHeader.Font.Size = 11
    lblHeader.Font.Bold = True
    lblSub.Font.Size = 8.5
    lblPrinterTitle.Font.Bold = True
    lblPrinterName.Font.Size = 8.5
    chkFitAllColumns.Font.Bold = True
    
    btnPreview.Font.Bold = True
    btnPreview.Font.Size = 9.5
    btnPrint.Font.Bold = True
    btnPrint.Font.Size = 9.5
    btnClose.Font.Size = 9.5
    
    btnAddOne.Font.Bold = True
    btnAddAll.Font.Bold = True
    btnRemoveOne.Font.Bold = True
    btnRemoveAll.Font.Bold = True
    btnMoveUp.Font.Bold = True
    btnMoveDown.Font.Bold = True
    btnApplyColor.Font.Bold = True
    
    chkAllVisible.Value = True
{captions_batchprint}

    ' Setup ComboBox Kho giay
    cboPaperSize.Clear
    cboPaperSize.AddItem UniText("{cbo_item0}")
    cboPaperSize.AddItem UniText("{cbo_item1}")
    cboPaperSize.AddItem UniText("{cbo_item2}")
    cboPaperSize.ListIndex = 0

    UpdatePrinterName
    LoadSheets
End Sub

Private Sub UpdatePrinterName()
    On Error Resume Next
    Dim pName As String
    pName = Application.ActivePrinter
    If Len(pName) > 42 Then pName = Left(pName, 40) & "..."
    lblPrinterName.Caption = pName
End Sub

Private Sub btnChangePrinter_Click()
    On Error Resume Next
    Application.Dialogs(xlDialogPrinterSetup).Show
    UpdatePrinterName
End Sub

Private Sub LoadSheets()
    lstSource.Clear
    lstTarget.Clear
    If ActiveWorkbook Is Nothing Then Exit Sub
    Dim ws As Object
    For Each ws In ActiveWorkbook.Sheets
        If chkAllVisible.Value Then
            If ws.Visible = -1 Then ' xlSheetVisible
                lstSource.AddItem ws.Name
            End If
        Else
            lstSource.AddItem ws.Name
        End If
    Next ws
    
    LoadTabColors
End Sub

Private Sub chkAllVisible_Click()
    LoadSheets
End Sub

Private Function GetTabColorName(ByVal c As Long) As String
    If c = -1 Then
        GetTabColorName = UniText("{esc('Không tô màu (No Color)')}")
        Exit Function
    End If
    Dim r As Long, g As Long, b As Long
    r = c And &HFF
    g = (c \\ &H100) And &HFF
    b = (c \\ &H10000) And &HFF
    
    If r > 180 And g < 80 And b < 80 Then
        GetTabColorName = UniText("{esc('Màu Đỏ (Red)')}")
    ElseIf r > 200 And g > 150 And b < 60 Then
        GetTabColorName = UniText("{esc('Màu Vàng (Yellow)')}")
    ElseIf r > 200 And g > 80 And g <= 150 And b < 60 Then
        GetTabColorName = UniText("{esc('Màu Cam (Orange)')}")
    ElseIf g > 140 And r < 120 And b < 120 Then
        GetTabColorName = UniText("{esc('Màu Xanh Lá (Green)')}")
    ElseIf b > 160 And r < 100 And g > 120 Then
        GetTabColorName = UniText("{esc('Màu Xanh Da Trời (Light Blue)')}")
    ElseIf b > 140 And r < 100 And g < 140 Then
        GetTabColorName = UniText("{esc('Màu Xanh Dương (Blue)')}")
    ElseIf r > 120 And b > 120 And g < 100 Then
        GetTabColorName = UniText("{esc('Màu Tím (Purple)')}")
    ElseIf r > 200 And g > 120 And b > 150 Then
        GetTabColorName = UniText("{esc('Màu Hồng (Pink)')}")
    ElseIf r = g And g = b Then
        GetTabColorName = UniText("{esc('Màu Xám (Gray)')}")
    Else
        GetTabColorName = "Màu RGB(" & r & "," & g & "," & b & ")"
    End If
End Function

Private Sub LoadTabColors()
    On Error Resume Next
    cboTabColor.Clear
    cboTabColor.ColumnCount = 2
    cboTabColor.ColumnWidths = "200 pt;0 pt"
    
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    Dim ws As Object
    Dim colDict As Object
    Set colDict = CreateObject("Scripting.Dictionary")
    
    For Each ws In ActiveWorkbook.Sheets
        If chkAllVisible.Value = False Or ws.Visible = -1 Then
            Dim cVal As Long
            If ws.Tab.ColorIndex = -4142 Then
                cVal = -1
            Else
                cVal = ws.Tab.Color
            End If
            
            If colDict.Exists(cVal) Then
                colDict(cVal) = colDict(cVal) + 1
            Else
                colDict.Add cVal, 1
            End If
        End If
    Next ws
    
    cboTabColor.AddItem UniText("{esc('-- Chọn màu Sheet để chuyển sang in --')}")
    cboTabColor.List(0, 1) = "-999"
    
    Dim k As Variant, cName As String, sCount As Long
    Dim hasColor As Boolean
    hasColor = False
    
    For Each k In colDict.Keys
        cVal = CLng(k)
        sCount = colDict(k)
        If cVal <> -1 Then
            hasColor = True
            cName = GetTabColorName(cVal)
            cboTabColor.AddItem cName & " (" & sCount & " sheet)"
            cboTabColor.List(cboTabColor.ListCount - 1, 1) = CStr(cVal)
        End If
    Next k
    
    If colDict.Exists(-1) Then
        cboTabColor.AddItem UniText("{esc('Không tô màu (No Color)')}") & " (" & colDict(-1) & " sheet)"
        cboTabColor.List(cboTabColor.ListCount - 1, 1) = "-1"
    End If
    
    If hasColor Or colDict.Count > 1 Then
        cboTabColor.ListIndex = 0
    Else
        cboTabColor.Clear
        cboTabColor.AddItem UniText("{esc('-- Workbook không có sheet nào tô màu --')}")
        cboTabColor.List(0, 1) = "-999"
        cboTabColor.ListIndex = 0
    End If
End Sub

Private Sub btnApplyColor_Click()
    On Error Resume Next
    If cboTabColor.ListIndex <= 0 Then
        MsgBoxUni "{esc('Vui lòng chọn một màu Tab trong danh sách!')}", vbExclamation
        Exit Sub
    End If
    
    Dim targetColor As Long
    targetColor = CLng(cboTabColor.List(cboTabColor.ListIndex, 1))
    If targetColor = -999 Then Exit Sub
    
    Dim i As Long, sName As String, ws As Object
    Dim matchColor As Long, movedCount As Long
    movedCount = 0
    
    For i = lstSource.ListCount - 1 To 0 Step -1
        sName = lstSource.List(i)
        Set ws = ActiveWorkbook.Sheets(sName)
        If Not ws Is Nothing Then
            If ws.Tab.ColorIndex = -4142 Then
                matchColor = -1
            Else
                matchColor = ws.Tab.Color
            End If
            
            If matchColor = targetColor Then
                lstTarget.AddItem sName
                lstSource.RemoveItem i
                movedCount = movedCount + 1
            End If
        End If
    Next i
    
    If movedCount > 0 Then
        MsgBoxUni UniText("{esc('Đã chuyển thành công')}") & " " & movedCount & " " & UniText("{esc('sheet có màu đã chọn sang danh sách in!')}"), vbInformation
    Else
        MsgBoxUni "{esc('Không tìm thấy sheet nào có màu này trong danh sách nguồn (hoặc đã chuyển hết)!')}", vbInformation
    End If
End Sub

Private Sub btnAddOne_Click()
    Dim i As Long
    For i = lstSource.ListCount - 1 To 0 Step -1
        If lstSource.Selected(i) Then
            lstTarget.AddItem lstSource.List(i)
            lstSource.RemoveItem i
        End If
    Next i
End Sub

Private Sub btnAddAll_Click()
    Dim i As Long
    For i = 0 To lstSource.ListCount - 1
        lstTarget.AddItem lstSource.List(i)
    Next i
    lstSource.Clear
End Sub

Private Sub btnRemoveOne_Click()
    Dim i As Long
    For i = lstTarget.ListCount - 1 To 0 Step -1
        If lstTarget.Selected(i) Then
            lstSource.AddItem lstTarget.List(i)
            lstTarget.RemoveItem i
        End If
    Next i
End Sub

Private Sub btnRemoveAll_Click()
    Dim i As Long
    For i = 0 To lstTarget.ListCount - 1
        lstSource.AddItem lstTarget.List(i)
    Next i
    lstTarget.Clear
End Sub

Private Sub lstSource_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnAddOne_Click
End Sub

Private Sub lstTarget_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnRemoveOne_Click
End Sub

Private Sub btnMoveUp_Click()
    Dim idx As Long, temp As String
    idx = lstTarget.ListIndex
    If idx > 0 Then
        temp = lstTarget.List(idx)
        lstTarget.List(idx) = lstTarget.List(idx - 1)
        lstTarget.List(idx - 1) = temp
        lstTarget.Selected(idx - 1) = True
    End If
End Sub

Private Sub btnMoveDown_Click()
    Dim idx As Long, temp As String
    idx = lstTarget.ListIndex
    If idx >= 0 And idx < lstTarget.ListCount - 1 Then
        temp = lstTarget.List(idx)
        lstTarget.List(idx) = lstTarget.List(idx + 1)
        lstTarget.List(idx + 1) = temp
        lstTarget.Selected(idx + 1) = True
    End If
End Sub

Private Sub ApplySettingsToSheet(ws As Object)
    On Error Resume Next
    ' 1. Kho giay
    If cboPaperSize.ListIndex = 1 Then
        ws.PageSetup.PaperSize = xlPaperA4
    ElseIf cboPaperSize.ListIndex = 2 Then
        ws.PageSetup.PaperSize = xlPaperA3
    End If

    ' 2. Fit All Columns on One Page
    If chkFitAllColumns.Value Then
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = False
    End If
End Sub

Private Sub btnPreview_Click()
    If lstTarget.ListCount = 0 Then
        MsgBoxUni "{esc('Vui lòng chọn ít nhất một sheet để xem trước!')}", vbExclamation
        Exit Sub
    End If

    Dim arr() As String, i As Long
    ReDim arr(0 To lstTarget.ListCount - 1)
    For i = 0 To lstTarget.ListCount - 1
        arr(i) = lstTarget.List(i)
        ApplySettingsToSheet ActiveWorkbook.Sheets(arr(i))
    Next i

    Dim actSheet As Object
    Set actSheet = ActiveSheet

    Me.Hide
    On Error Resume Next
    ActiveWorkbook.Sheets(arr).Select
    ActiveWindow.SelectedSheets.PrintPreview

    If Not actSheet Is Nothing Then actSheet.Select
    Me.Show vbModeless
End Sub

Private Sub btnPrint_Click()
    If lstTarget.ListCount = 0 Then
        MsgBoxUni "{esc('Vui lòng chọn ít nhất một sheet để in!')}", vbExclamation
        Exit Sub
    End If

    Dim pInfo As String
    pInfo = Application.ActivePrinter
    If Len(pInfo) > 35 Then pInfo = Left(pInfo, 32) & "..."

    Dim msg As String
    msg = UniText("{esc('Bạn có chắc chắn muốn in trực tiếp')}") & " " & lstTarget.ListCount & " " & UniText("{esc('sheet đã chọn ra máy in:')}") & vbCrLf & _
          "[" & pInfo & "] " & UniText("{esc('không?')}")

    If MsgBoxUni(msg, vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    Dim actSheet As Object
    Set actSheet = ActiveSheet
    Me.Hide

    Dim i As Long, sName As String
    For i = 0 To lstTarget.ListCount - 1
        sName = lstTarget.List(i)
        ApplySettingsToSheet ActiveWorkbook.Sheets(sName)
        On Error Resume Next
        ActiveWorkbook.Sheets(sName).PrintOut
    Next i

    If Not actSheet Is Nothing Then actSheet.Select
    Me.Show vbModeless
    MsgBoxUni UniText("{esc('Đã gửi thành công')}") & " " & lstTarget.ListCount & " " & UniText("{esc('sheet tới máy in!')}"), vbInformation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
'''


# Form 3: frmNumberSeparator Code
code_frmNumberSeparator = f"""Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Function UniText(ByVal s As String) As String
    Dim pos As Long, hexStr As String, res As String
    res = s
    pos = InStr(res, "\\u")
    Do While pos > 0
        hexStr = Mid(res, pos + 2, 4)
        res = Left(res, pos - 1) & ChrW(CLng("&H" & hexStr)) & Mid(res, pos + 6)
        pos = InStr(res, "\\u")
    Loop
    UniText = res
End Function

Private Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    If Len(title) = 0 Then title = UniText("{esc('AVTool - Đổi Dấu Phân Cách Số')}")
    MsgBoxUni = MessageBoxW(0, StrPtr(UniText(prompt)), StrPtr(title), buttons)
End Function

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim c As Control
    For Each c In Me.Controls
        c.Font.Name = "Segoe UI"
        c.Font.Size = 9
    Next c
    lblHeader.Font.Size = 10.5
    lblHeader.Font.Bold = True
    lblSub.Font.Size = 8.5
    optVN.Font.Bold = True
    optUS.Font.Bold = True
    btnApply.Font.Bold = True
    btnApply.Font.Size = 9.5
    btnClose.Font.Size = 9.5
    
    Me.Caption = "AVTool - Doi Dau Phan Cach So"
    lblHeader.Caption = UniText("{esc('ĐỔI DẤU PHÂN CÁCH SỐ')}")
    lblSub.Caption = UniText("{esc('Chuyển đổi dấu thập phân & hàng nghìn tức thì trong Excel')}")
    fraOptions.Caption = UniText("{esc(' Tùy chọn định dạng số ')}")
    optVN.Caption = UniText("{esc('Chuẩn Việt Nam:  1.234.567,89  (Phẩy thập phân, Chấm hàng nghìn)')}")
    optUS.Caption = UniText("{esc('Chuẩn Quốc tế/US: 1,234,567.89 (Chấm thập phân, Phẩy hàng nghìn)')}")
    optSystem.Caption = UniText("{esc('Theo cài đặt Windows (Use system separators)')}")
    btnApply.Caption = UniText("{esc('Áp dụng')}")
    btnClose.Caption = UniText("{esc('Đóng')}")
    
    RefreshStatus
End Sub

Private Sub RefreshStatus()
    On Error Resume Next
    If Application.UseSystemSeparators Then
        optSystem.Value = True
        lblCurrent.Caption = UniText("{esc('Hiện tại: Theo cài đặt Windows (Thập phân: ')}") & Application.International(xlDecimalSeparator) & UniText("{esc(', Hàng nghìn: ')}") & Application.International(xlThousandsSeparator) & ")"
    Else
        If Application.DecimalSeparator = "," And Application.ThousandsSeparator = "." Then
            optVN.Value = True
            lblCurrent.Caption = UniText("{esc('Hiện tại: Chuẩn Việt Nam (Thập phân: phẩy \',\', Hàng nghìn: chấm \'.\')')}")
        ElseIf Application.DecimalSeparator = "." And Application.ThousandsSeparator = "," Then
            optUS.Value = True
            lblCurrent.Caption = UniText("{esc('Hiện tại: Chuẩn Quốc tế / US (Thập phân: chấm \'.\', Hàng nghìn: phẩy \',\')')}")
        Else
            lblCurrent.Caption = UniText("{esc('Hiện tại: Tùy chỉnh (Thập phân: ')}") & Application.DecimalSeparator & UniText("{esc(', Hàng nghìn: ')}") & Application.ThousandsSeparator & ")"
        End If
    End If
End Sub

Private Sub btnApply_Click()
    On Error GoTo ErrHandle
    If optVN.Value Then
        Application.UseSystemSeparators = False
        Application.DecimalSeparator = ","
        Application.ThousandsSeparator = "."
        MsgBoxUni "{esc('Đã chuyển sang Chuẩn Việt Nam thành công!')}" & vbCrLf & "{esc('(Thập phân: dấu phẩy \',\' | Hàng nghìn: dấu chấm \'.\')')}", vbInformation
    ElseIf optUS.Value Then
        Application.UseSystemSeparators = False
        Application.DecimalSeparator = "."
        Application.ThousandsSeparator = ","
        MsgBoxUni "{esc('Đã chuyển sang Chuẩn Quốc tế / US thành công!')}" & vbCrLf & "{esc('(Thập phân: dấu chấm \'.\' | Hàng nghìn: dấu phẩy \',\')')}", vbInformation
    ElseIf optSystem.Value Then
        Application.UseSystemSeparators = True
        MsgBoxUni "{esc('Đã chuyển về Sử dụng dấu theo Hệ điều hành Windows (Use system separators)!')}", vbInformation
    End If
    RefreshStatus
    Exit Sub
ErrHandle:
    MsgBoxUni "{esc('Lỗi khi thiết lập dấu phân cách: ')}" & Err.Description, vbCritical
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
"""


# Form 4: frmFixDValue Code
code_frmFixDValue = f"""Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Function UniText(ByVal s As String) As String
    Dim pos As Long, hexStr As String, res As String
    res = s
    pos = InStr(res, "\\u")
    Do While pos > 0
        hexStr = Mid(res, pos + 2, 4)
        res = Left(res, pos - 1) & ChrW(CLng("&H" & hexStr)) & Mid(res, pos + 6)
        pos = InStr(res, "\\u")
    Loop
    UniText = res
End Function

Private Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    If Len(title) = 0 Then title = UniText("{esc('AVTool - Sửa Lỗi DValue')}")
    MsgBoxUni = MessageBoxW(0, StrPtr(UniText(prompt)), StrPtr(title), buttons)
End Function

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim c As Control
    For Each c In Me.Controls
        c.Font.Name = "Segoe UI"
        c.Font.Size = 9
    Next c
    lblHeader.Font.Size = 10.5
    lblHeader.Font.Bold = True
    lblSub.Font.Size = 8.5
    lblStatusSummary.Font.Bold = True
    btnFix.Font.Bold = True
    btnFix.Font.Size = 9.5
    btnRescan.Font.Size = 9
    btnClose.Font.Size = 9.5
    
    Me.Caption = "AVTool - Sua Loi Duong Dan DValue & Cong Thuc"
    lblHeader.Caption = UniText("{esc('SỬA LỖI ĐƯỜNG DẪN DVALUE (#NAME?)')}")
    lblSub.Caption = UniText("{esc('Tự động cắt bỏ đường dẫn file lạ, khôi phục công thức khối lượng DValue tức thì')}")
    
    fraStatus.Caption = UniText("{esc(' Tình trạng quét trong Workbook ')}")
    fraScope.Caption = UniText("{esc(' Phạm vi thực hiện ')}")
    fraOptions.Caption = UniText("{esc(' Tùy chọn sửa lỗi ')}")
    
    optScopeSheet.Value = True
    If Not ActiveSheet Is Nothing Then
        optScopeSheet.Caption = UniText("{esc('Sheet hiện tại: ')}") & ActiveSheet.Name
    Else
        optScopeSheet.Caption = UniText("{esc('Sheet hiện tại')}")
    End If
    optScopeWorkbook.Caption = UniText("{esc('Toàn bộ các Sheet trong File (Workbook)')}")
    optScopeSelection.Caption = UniText("{esc('Chỉ vùng ô đang chọn (Selection)')}")
    
    chkCleanPath.Caption = UniText("{esc('Xóa đường dẫn ngoài dính vào DValue (Đưa về =DValue())')}")
    chkCleanPath.Value = True
    
    chkNormalizeOthers.Caption = UniText("{esc('Chuẩn hóa các hàm khối lượng cũ khác (=Tinh, =Eval, =DienGiai)')}")
    chkNormalizeOthers.Value = False
    
    chkRecalculate.Caption = UniText("{esc('Tự động tính toán lại toàn bộ sau khi sửa (CalculateFullRebuild)')}")
    chkRecalculate.Value = True
    
    btnRescan.Caption = UniText("{esc('Quét lại')}")
    btnFix.Caption = UniText("{esc('⚡ SỬA LỖI NGAY')}")
    btnClose.Caption = UniText("{esc('Đóng')}")
    
    ScanErrors
End Sub

Private Sub ScanErrors()
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        lblStatusSummary.Caption = UniText("{esc('Chưa mở file Excel nào để quét!')}")
        txtDetails.Text = ""
        Exit Sub
    End If
    
    Dim totalErrDValue As Long, totalSheets As Long
    Dim sheetList As String, sampleFormula As String
    totalErrDValue = 0
    totalSheets = 0
    sheetList = ""
    sampleFormula = ""
    
    Dim ws As Worksheet, fCells As Range, c As Range, f As String
    For Each ws In wb.Worksheets
        Set fCells = Nothing
        Set fCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Not fCells Is Nothing Then
            Dim countInSheet As Long
            countInSheet = 0
            For Each c In fCells
                f = c.Formula
                If (InStr(1, f, "!DValue(", vbTextCompare) > 0 Or InStr(1, f, "!dvalue(", vbTextCompare) > 0) And (InStr(1, f, ":\\") > 0 Or InStr(1, f, "'.*!'") > 0 Or InStr(1, f, "[") > 0 Or InStr(1, f, ".xla") > 0) Then
                    countInSheet = countInSheet + 1
                    If Len(sampleFormula) = 0 Then sampleFormula = f
                End If
            Next c
            If countInSheet > 0 Then
                totalSheets = totalSheets + 1
                totalErrDValue = totalErrDValue + countInSheet
                sheetList = sheetList & "• Sheet '" & ws.Name & "': " & countInSheet & " " & UniText("{esc('ô')}") & vbCrLf
            End If
        End If
    Next ws
    
    If totalErrDValue > 0 Then
        lblStatusSummary.Caption = UniText("{esc('Phát hiện ')}") & totalErrDValue & UniText("{esc(' ô dính đường dẫn ngoài DValue trên ')}") & totalSheets & UniText("{esc(' Sheet!')}")
        lblStatusSummary.ForeColor = RGB(180, 40, 0)
        
        Dim detailStr As String
        detailStr = sheetList
        If Len(sampleFormula) > 0 Then
            detailStr = detailStr & vbCrLf & UniText("{esc('Ví dụ mẫu: ')}") & sampleFormula & " -> =DValue(...)"
        End If
        txtDetails.Text = detailStr
    Else
        lblStatusSummary.Caption = UniText("{esc('Không tìm thấy ô nào bị dính đường dẫn ngoài DValue.')}")
        lblStatusSummary.ForeColor = RGB(0, 128, 0)
        txtDetails.Text = UniText("{esc('Bảng tính hiện tại không bị lỗi đường dẫn DValue. Bạn vẫn có thể bấm Sửa Lỗi Ngay nếu muốn kiểm tra hoặc chuẩn hóa các hàm khác.')}")
    End If
End Sub

Private Sub btnRescan_Click()
    ScanErrors
End Sub

Private Sub btnFix_Click()
    On Error GoTo ErrHandle
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBoxUni "{esc('Vui lòng mở một file Excel trước khi thực hiện!')}", vbExclamation
        Exit Sub
    End If
    
    Dim targetScope As Integer
    If optScopeSheet.Value Then
        targetScope = 1
    ElseIf optScopeWorkbook.Value Then
        targetScope = 2
    Else
        targetScope = 3
    End If
    
    Dim bCleanPath As Boolean, bNormalizeOthers As Boolean, bRecalc As Boolean
    bCleanPath = chkCleanPath.Value
    bNormalizeOthers = chkNormalizeOthers.Value
    bRecalc = chkRecalculate.Value
    
    If Not bCleanPath And Not bNormalizeOthers Then
        MsgBoxUni "{esc('Vui lòng chọn ít nhất một tùy chọn sửa lỗi!')}", vbExclamation
        Exit Sub
    End If
    
    Dim targetSheets As New Collection
    Select Case targetScope
        Case 1
            targetSheets.Add ActiveSheet
        Case 2
            Dim s As Worksheet
            For Each s In wb.Worksheets
                targetSheets.Add s
            Next s
        Case 3
            targetSheets.Add ActiveSheet
    End Select
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim ws As Worksheet, rng As Range
    Dim otherFuncs As Variant
    otherFuncs = Array("Tinh", "Eval", "DienGiai", "CDTDValue", "TinhToan", "KL", "KLDT", "Cal", "Evaluate")
    Dim ofName As Variant
    
    For Each ws In targetSheets
        If targetScope = 3 Then
            Set rng = Selection
        Else
            Set rng = ws.UsedRange
        End If
        
        If Not rng Is Nothing Then
            If bCleanPath Then
                rng.Replace What:="'*!DValue(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
                rng.Replace What:="'*!dvalue(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
                rng.Replace What:="[*]!DValue(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
                rng.Replace What:="[*]!dvalue(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
                rng.Replace What:="*!DValue(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
                rng.Replace What:="*!dvalue(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
            End If
            
            If bNormalizeOthers Then
                For Each ofName In otherFuncs
                    rng.Replace What:="'*!" & CStr(ofName) & "(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
                    rng.Replace What:="[*]!" & CStr(ofName) & "(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
                    rng.Replace What:="*!" & CStr(ofName) & "(", Replacement:="DValue(", LookAt:=xlPart, MatchCase:=False
                    rng.Replace What:="=" & CStr(ofName) & "(", Replacement:="=DValue(", LookAt:=xlPart, MatchCase:=False
                Next ofName
            End If
        End If
    Next ws
    
    If bRecalc Then
        On Error Resume Next
        Application.CalculateFullRebuild
        If Err.Number <> 0 Then Application.Calculate
        On Error GoTo ErrHandle
    End If
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ScanErrors
    MsgBoxUni "{esc('ĐÃ SỬA LỖI THÀNH CÔNG!')}" & vbCrLf & vbCrLf & _
              "{esc('- Toàn bộ đường dẫn ngoài đã được loại bỏ.')}" & vbCrLf & _
              "{esc('- Công thức đã được đưa về =DValue() chuẩn của AVTool.')}" & vbCrLf & _
              "{esc('- Toàn bộ bảng tính đã được tính toán lại (Recalculated) tức thì!')}", vbInformation
    Exit Sub
    
ErrHandle:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBoxUni "{esc('Có lỗi xảy ra trong quá trình sửa lỗi: ')}" & Err.Description, vbCritical
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
"""


# =============================================================================
# Form 5: frmNoiSuyMatran (Chuan Unicode 100%)
# =============================================================================
code_frmNoiSuyMatran = r"""#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
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

Private Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "AVTool") As Long
    MsgBoxUni = MessageBoxW(0, StrPtr(UniText(prompt)), StrPtr(UniText(title)), buttons)
End Function

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim c As Control
    For Each c In Me.Controls
        c.Font.Name = "Segoe UI"
        c.Font.Size = 9
    Next c
    lblHeader.Font.Bold = True
    lblHeader.Font.Size = 10.5
    lblSub.Font.Size = 8.5
    lblPreview.Font.Bold = True
    btnInsert.Font.Bold = True
    btnInsert.Font.Size = 9.5
    chkAbsolute.Value = True
    
    Me.Caption = "AVTool_NoiSuyMatran_Win"
    lblHeader.Caption = UniText("N\u1ED8I SUY MA TR\u1EACN 2 CHI\u1EC0U (B\u1EA2NG TRA 2 TR\u1EE4C X & Y)")
    lblSub.Caption = UniText("Tra c\u1EE9u song tuy\u1EBFn t\u00EDnh (Bilinear). T\u1EF1 \u0111\u1ED9ng kh\u1EDBp k\u00EDch th\u01B0\u1EDBc & sinh c\u00F4ng th\u1EE9c")
    fraTable.Caption = UniText(" Thi\u1EBFt l\u1EADp v\u00F9ng b\u1EA3ng tra ma tr\u1EADn ")
    lblRangeX.Caption = UniText("H\u00E0ng m\u1ED1c X (Ngang):")
    lblRangeY.Caption = UniText("C\u1ED9t m\u1ED1c Y (D\u1ECDc):")
    lblRangeGT.Caption = UniText("Ma tr\u1EADn s\u1ED1 li\u1EC7u (GT):")
    lblDimCheck.Caption = UniText("Vui l\u00F2ng qu\u00E9t ch\u1ECDn c\u00E1c v\u00F9ng tr\u00EAn b\u1EA3ng t\u00EDnh")
    fraQuery.Caption = UniText(" Gi\u00E1 tr\u1ECB c\u1EA7n tra c\u1EE9u & \u00D4 \u0111\u1EB7t k\u1EBFt qu\u1EA3 ")
    lblValX.Caption = UniText("M\u1ED1c Xu:")
    lblValY.Caption = UniText("M\u1ED1c Yu:")
    lblDest.Caption = UniText("\u00D4 k\u1EBFt qu\u1EA3:")
    chkAbsolute.Caption = UniText("C\u1ED1 \u0111\u1ECBnh b\u1EA3ng ($)")
    fraPreview.Caption = UniText(" Xem tr\u01B0\u1EDBc k\u1EBFt qu\u1EA3 & C\u00F4ng th\u1EE9c sinh ra ")
    lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: (Ch\u1EDD nh\u1EADp th\u00F4ng tin...)")
    btnSample.Caption = UniText("\uD83D\uDCCB T\u1EA1o B\u1EA3ng M\u1EABu")
    btnInsert.Caption = UniText("\u26A1 CH\u00C8N C\u00D4NG TH\u1EE8C")
    btnClose.Caption = UniText("\u0110\u00F3ng")
    
    If Not ActiveCell Is Nothing Then
        txtDest.Text = ActiveCell.Address(False, False)
    End If
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    #If VBA7 Then
        Dim hWndForm As LongPtr
    #Else
        Dim hWndForm As Long
    #End If
    hWndForm = FindWindowA("ThunderDFrame", "AVTool_NoiSuyMatran_Win")
    If hWndForm <> 0 Then
        SetWindowTextW hWndForm, StrPtr(UniText("AVTool - N\u1ED9i Suy Ma Tr\u1EADn 2 Chi\u1EC1u (2D Bilinear)"))
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPickX_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn H\u00E0ng m\u1ED1c X (Tr\u1EE5c ngang):"), UniText("Ch\u1ECDn H\u00E0ng M\u1ED1c X"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeX.Text = rng.Address(True, True)
        Else
            txtRangeX.Text = rng.Address(False, False)
        End If
        UpdatePreview
    End If
End Sub

Private Sub btnPickY_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn C\u1ED9t m\u1ED1c Y (Tr\u1EE5c d\u1ECDc):"), UniText("Ch\u1ECDn C\u1ED9t M\u1ED1c Y"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeY.Text = rng.Address(True, True)
        Else
            txtRangeY.Text = rng.Address(False, False)
        End If
        UpdatePreview
    End If
End Sub

Private Sub btnPickGT_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn v\u00F9ng Ma tr\u1EADn s\u1ED1 li\u1EC7u (GT):"), UniText("Ch\u1ECDn V\u00F9ng Ma Tr\u1EADn"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeGT.Text = rng.Address(True, True)
        Else
            txtRangeGT.Text = rng.Address(False, False)
        End If
        UpdatePreview
    End If
End Sub

Private Sub btnPickValX_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn \u00F4 ch\u1EE9a gi\u00E1 tr\u1ECB m\u1ED1c Xu:"), UniText("Ch\u1ECDn \u00D4 Gi\u00E1 Tr\u1ECB Xu"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        txtValX.Text = rng.Address(False, False)
        UpdatePreview
    End If
End Sub

Private Sub btnPickValY_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn \u00F4 ch\u1EE9a gi\u00E1 tr\u1ECB m\u1ED1c Yu:"), UniText("Ch\u1ECDn \u00D4 Gi\u00E1 Tr\u1ECB Yu"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        txtValY.Text = rng.Address(False, False)
        UpdatePreview
    End If
End Sub

Private Sub btnPickDest_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn \u00F4 \u0111\u1EB7t c\u00F4ng th\u1EE9c k\u1EBFt qu\u1EA3:"), UniText("Ch\u1ECDn \u00D4 K\u1EBFt Qu\u1EA3"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        txtDest.Text = rng.Address(False, False)
    End If
End Sub

Private Sub txtRangeX_Change()
    UpdatePreview
End Sub

Private Sub txtRangeY_Change()
    UpdatePreview
End Sub

Private Sub txtRangeGT_Change()
    UpdatePreview
End Sub

Private Sub txtValX_Change()
    UpdatePreview
End Sub

Private Sub txtValY_Change()
    UpdatePreview
End Sub

Private Sub chkAbsolute_Click()
    Dim rX As Range, rY As Range, rGT As Range
    On Error Resume Next
    Set rX = Range(txtRangeX.Text)
    Set rY = Range(txtRangeY.Text)
    Set rGT = Range(txtRangeGT.Text)
    If Not rX Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeX.Text = rX.Address(True, True)
        Else
            txtRangeX.Text = rX.Address(False, False)
        End If
    End If
    If Not rY Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeY.Text = rY.Address(True, True)
        Else
            txtRangeY.Text = rY.Address(False, False)
        End If
    End If
    If Not rGT Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeGT.Text = rGT.Address(True, True)
        Else
            txtRangeGT.Text = rGT.Address(False, False)
        End If
    End If
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    Dim rX As Range, rY As Range, rGT As Range
    Dim vX As Double, vY As Double
    Dim res As Variant
    Dim formStr As String
    Dim sX As String, sY As String
    
    On Error Resume Next
    Set rX = Range(txtRangeX.Text)
    Set rY = Range(txtRangeY.Text)
    Set rGT = Range(txtRangeGT.Text)
    
    If rX Is Nothing Or rY Is Nothing Or rGT Is Nothing Then
        lblDimCheck.Caption = UniText("Ch\u01B0a ch\u1ECDn \u0111\u1EE7 3 v\u00F9ng (X, Y, GT)")
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: (Ch\u01B0a \u0111\u1EE7 d\u1EEF li\u1EC7u)")
        txtFormula.Text = ""
        Exit Sub
    End If
    
    If rGT.Columns.Count <> rX.Count Or rGT.Rows.Count <> rY.Count Then
        lblDimCheck.Caption = UniText("\u274C Sai kh\u00E1c: GT (" & rGT.Rows.Count & "x" & rGT.Columns.Count & ") kh\u00E1c X(" & rX.Count & "), Y(" & rY.Count & ")")
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: \u274C Sai l\u1EC7ch k\u00EDch th\u01B0\u1EDBc")
        txtFormula.Text = ""
        Exit Sub
    Else
        lblDimCheck.Caption = UniText("\u2713 \u0110\u00E3 kh\u1EDBp: " & rX.Count & " c\u1ED9t X \u00D7 " & rY.Count & " d\u00F2ng Y")
    End If
    
    sX = Trim(txtValX.Text)
    sY = Trim(txtValY.Text)
    If Len(sX) = 0 Or Len(sY) = 0 Then
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3: (Ch\u1EDD nh\u1EADp Xu, Yu)")
        txtFormula.Text = "=NoisuyMatran(" & txtRangeX.Text & ", " & txtRangeY.Text & ", " & txtRangeGT.Text & ", Xu, Yu)"
        Exit Sub
    End If
    
    If IsNumeric(sX) Then
        vX = CDbl(sX)
        sX = Replace(sX, ",", ".")
    Else
        vX = Val(Replace(CStr(Range(sX).Value), ",", "."))
    End If
    If IsNumeric(sY) Then
        vY = CDbl(sY)
        sY = Replace(sY, ",", ".")
    Else
        vY = Val(Replace(CStr(Range(sY).Value), ",", "."))
    End If
    
    formStr = "=NoisuyMatran(" & txtRangeX.Text & ", " & txtRangeY.Text & ", " & txtRangeGT.Text & ", " & sX & ", " & sY & ")"
    txtFormula.Text = formStr
    
    res = NoisuyMatran(rX, rY, rGT, vX, vY)
    If IsError(res) Or VarType(res) = vbError Then
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: \u274C L\u1ED7i t\u00EDnh to\u00E1n (#VALUE!)")
    Else
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: " & Format(res, "#,##0.0000"))
    End If
End Sub

Private Sub btnInsert_Click()
    Dim destRng As Range
    On Error Resume Next
    If Len(Trim(txtFormula.Text)) = 0 Then
        MsgBoxUni "Vui l\u00F2ng nh\u1EADp \u0111\u1EA7y \u0111\u1EE7 c\u00E1c th\u00F4ng s\u1ED1 tr\u01B0\u1EDBc khi ch\u00E8n!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    If Len(Trim(txtDest.Text)) > 0 Then
        Set destRng = Range(txtDest.Text)
    Else
        Set destRng = ActiveCell
    End If
    
    If destRng Is Nothing Then
        MsgBoxUni "Ch\u01B0a ch\u1ECDn \u00F4 \u0111\u1EB7t k\u1EBFt qu\u1EA3!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    InsertNoiSuyToActiveWorkbook showMsg:=False
    destRng.Formula = txtFormula.Text
    destRng.Select
    
    MsgBoxUni "\u0110\u00C3 CH\u00C8N C\u00D4NG TH\u1EE8C V\u00C0O \u00D4 " & destRng.Address(False, False) & " TH\u00C0NH C\u00D4NG!" & vbCrLf & vbCrLf & _
              "C\u00F4ng th\u1EE9c: " & txtFormula.Text, vbInformation, "AVTool"
End Sub

Private Sub btnSample_Click()
    CreateSampleTableAll True
End Sub
"""

code_frmNoiSuyDay = r"""#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
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

Private Function MsgBoxUni(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "AVTool") As Long
    MsgBoxUni = MessageBoxW(0, StrPtr(UniText(prompt)), StrPtr(UniText(title)), buttons)
End Function

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim c As Control
    For Each c In Me.Controls
        c.Font.Name = "Segoe UI"
        c.Font.Size = 9
    Next c
    lblHeader.Font.Bold = True
    lblHeader.Font.Size = 10.5
    lblSub.Font.Size = 8.5
    lblPreview.Font.Bold = True
    btnInsert.Font.Bold = True
    btnInsert.Font.Size = 9.5
    chkAbsolute.Value = True
    
    Me.Caption = "AVTool_NoiSuyDay_Win"
    lblHeader.Caption = UniText("N\u1ED8I SUY D\u00C3Y S\u1ED0 1 CHI\u1EC0U (TUY\u1EBE N T\u00CDNH)")
    lblSub.Caption = UniText("Tra c\u1EE9u tuy\u1EBFn t\u00EDnh theo d\u00E3y h\u00E0ng ho\u1EB7c c\u1ED9t. T\u1EF1 \u0111\u1ED9ng s\u1EAFp x\u1EBFp m\u1ED1c t\u0103ng d\u1EA7n")
    fraTable.Caption = UniText(" Thi\u1EBFt l\u1EADp 2 d\u00E3y d\u1EEF li\u1EC7u b\u1EA3ng tra ")
    lblRangeX.Caption = UniText("D\u00E3y m\u1ED1c tra (D\u00E3y X):")
    lblRangeY.Caption = UniText("D\u00E3y gi\u00E1 tr\u1ECB (D\u00E3y Y):")
    lblDimCheck.Caption = UniText("Vui l\u00F2ng qu\u00E9t ch\u1ECDn d\u00E3y X v\u00E0 d\u00E3y Y tr\u00EAn b\u1EA3ng t\u00EDnh")
    fraQuery.Caption = UniText(" Gi\u00E1 tr\u1ECB c\u1EA7n tra c\u1EE9u & \u00D4 \u0111\u1EB7t k\u1EBFt qu\u1EA3 ")
    lblValX.Caption = UniText("Gi\u00E1 tr\u1ECB Xu:")
    lblDest.Caption = UniText("\u00D4 k\u1EBFt qu\u1EA3:")
    chkAbsolute.Caption = UniText("C\u1ED1 \u0111\u1ECBnh d\u00E3y ($)")
    fraPreview.Caption = UniText(" Xem tr\u01B0\u1EDBc k\u1EBFt qu\u1EA3 & C\u00F4ng th\u1EE9c sinh ra ")
    lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: (Ch\u1EDD nh\u1EADp th\u00F4ng tin...)")
    btnSample.Caption = UniText("\uD83D\uDCCB T\u1EA1o D\u00E3y M\u1EABu")
    btnInsert.Caption = UniText("\u26A1 CH\u00C8N C\u00D4NG TH\u1EE8C")
    btnClose.Caption = UniText("\u0110\u00F3ng")
    
    If Not ActiveCell Is Nothing Then
        txtDest.Text = ActiveCell.Address(False, False)
    End If
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    #If VBA7 Then
        Dim hWndForm As LongPtr
    #Else
        Dim hWndForm As Long
    #End If
    hWndForm = FindWindowA("ThunderDFrame", "AVTool_NoiSuyDay_Win")
    If hWndForm <> 0 Then
        SetWindowTextW hWndForm, StrPtr(UniText("AVTool - N\u1ED9i Suy D\u00E3y S\u1ED1 1 Chi\u1EC1u (1D Linear)"))
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPickX_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn d\u00E3y m\u1ED1c tra (D\u00E3y X):"), UniText("Ch\u1ECDn D\u00E3y M\u1ED1c X"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeX.Text = rng.Address(True, True)
        Else
            txtRangeX.Text = rng.Address(False, False)
        End If
        UpdatePreview
    End If
End Sub

Private Sub btnPickY_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn d\u00E3y gi\u00E1 tr\u1ECB tra (D\u00E3y Y):"), UniText("Ch\u1ECDn D\u00E3y Gi\u00E1 Tr\u1ECB Y"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeY.Text = rng.Address(True, True)
        Else
            txtRangeY.Text = rng.Address(False, False)
        End If
        UpdatePreview
    End If
End Sub

Private Sub btnPickValX_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn \u00F4 ch\u1EE9a m\u1ED1c c\u1EA7n tra Xu:"), UniText("Ch\u1ECDn \u00D4 M\u1ED1c Xu"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        txtValX.Text = rng.Address(False, False)
        UpdatePreview
    End If
End Sub

Private Sub btnPickDest_Click()
    Dim rng As Range
    Me.Hide
    On Error Resume Next
    Set rng = Application.InputBox(UniText("Ch\u1ECDn \u00F4 \u0111\u1EB7t c\u00F4ng th\u1EE9c k\u1EBFt qu\u1EA3:"), UniText("Ch\u1ECDn \u00D4 K\u1EBFt Qu\u1EA3"), Type:=8)
    On Error GoTo 0
    Me.Show vbModeless
    If Not rng Is Nothing Then
        txtDest.Text = rng.Address(False, False)
    End If
End Sub

Private Sub txtRangeX_Change()
    UpdatePreview
End Sub

Private Sub txtRangeY_Change()
    UpdatePreview
End Sub

Private Sub txtValX_Change()
    UpdatePreview
End Sub

Private Sub chkAbsolute_Click()
    Dim rX As Range, rY As Range
    On Error Resume Next
    Set rX = Range(txtRangeX.Text)
    Set rY = Range(txtRangeY.Text)
    If Not rX Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeX.Text = rX.Address(True, True)
        Else
            txtRangeX.Text = rX.Address(False, False)
        End If
    End If
    If Not rY Is Nothing Then
        If chkAbsolute.Value Then
            txtRangeY.Text = rY.Address(True, True)
        Else
            txtRangeY.Text = rY.Address(False, False)
        End If
    End If
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    Dim rX As Range, rY As Range
    Dim vX As Double
    Dim res As Variant
    Dim formStr As String
    Dim sX As String
    
    On Error Resume Next
    Set rX = Range(txtRangeX.Text)
    Set rY = Range(txtRangeY.Text)
    
    If rX Is Nothing Or rY Is Nothing Then
        lblDimCheck.Caption = UniText("Ch\u01B0a ch\u1ECDn \u0111\u1EE7 2 d\u00E3y (X v\u00E0 Y)")
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: (Ch\u01B0a \u0111\u1EE7 d\u1EEF li\u1EC7u)")
        txtFormula.Text = ""
        Exit Sub
    End If
    
    If rX.Count <> rY.Count Then
        lblDimCheck.Caption = UniText("\u274C S\u1ED1 ph\u1EA7n t\u1EED kh\u00F4ng kh\u1EDBp: X c\u00F3 " & rX.Count & " \u00F4, Y c\u00F3 " & rY.Count & " \u00F4")
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: \u274C Sai kh\u00E1c s\u1ED1 ph\u1EA7n t\u1EED")
        txtFormula.Text = ""
        Exit Sub
    Else
        lblDimCheck.Caption = UniText("\u2713 \u0110\u00E3 kh\u1EDBp: " & rX.Count & " ph\u1EA7n t\u1EED")
    End If
    
    sX = Trim(txtValX.Text)
    If Len(sX) = 0 Then
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3: (Ch\u1EDD nh\u1EADp Xu)")
        txtFormula.Text = "=NoisuyDay(" & txtRangeX.Text & ", " & txtRangeY.Text & ", Xu)"
        Exit Sub
    End If
    
    If IsNumeric(sX) Then
        vX = CDbl(sX)
        sX = Replace(sX, ",", ".")
    Else
        vX = Val(Replace(CStr(Range(sX).Value), ",", "."))
    End If
    
    formStr = "=NoisuyDay(" & txtRangeX.Text & ", " & txtRangeY.Text & ", " & sX & ")"
    txtFormula.Text = formStr
    
    res = NoisuyDay(rX, rY, vX)
    If IsError(res) Or VarType(res) = vbError Then
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: \u274C L\u1ED7i t\u00EDnh to\u00E1n (#VALUE!)")
    Else
        lblPreview.Caption = UniText("K\u1EBFt qu\u1EA3 xem tr\u01B0\u1EDBc: " & Format(res, "#,##0.0000"))
    End If
End Sub

Private Sub btnInsert_Click()
    Dim destRng As Range
    On Error Resume Next
    If Len(Trim(txtFormula.Text)) = 0 Then
        MsgBoxUni "Vui l\u00F2ng nh\u1EADp \u0111\u1EA7y \u0111\u1EE7 c\u00E1c th\u00F4ng s\u1ED1 tr\u01B0\u1EDBc khi ch\u00E8n!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    If Len(Trim(txtDest.Text)) > 0 Then
        Set destRng = Range(txtDest.Text)
    Else
        Set destRng = ActiveCell
    End If
    
    If destRng Is Nothing Then
        MsgBoxUni "Ch\u01B0a ch\u1ECDn \u00F4 \u0111\u1EB7t k\u1EBFt qu\u1EA3!", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    InsertNoiSuyToActiveWorkbook showMsg:=False
    destRng.Formula = txtFormula.Text
    destRng.Select
    
    MsgBoxUni "\u0110\u00C3 CH\u00C8N C\u00D4NG TH\u1EE8C V\u00C0O \u00D4 " & destRng.Address(False, False) & " TH\u00C0NH C\u00D4NG!" & vbCrLf & vbCrLf & _
              "C\u00F4ng th\u1EE9c: " & txtFormula.Text, vbInformation, "AVTool"
End Sub

Private Sub btnSample_Click()
    CreateSampleTableAll True
End Sub
"""

ps_script = f"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $wb = $excel.Workbooks.Add()

    # Them Sheet _IntelliSense_ nhac lenh trang tinh chuan Excel-DNA cho tat ca cac ham noi suy
    $wsIntel = $wb.Worksheets.Add()
    $wsIntel.Name = "_IntelliSense_"
    $wsIntel.Range("A1").Value = "FunctionInfo"
    $wsIntel.Range("B1").Value = 1

    # 1. Noisuy2b
    $wsIntel.Range("A2").Value = "Noisuy2b"
    $wsIntel.Range("B2").Value = "Ham noi suy tuyen tinh 1 chieu giua 2 diem kep (Xs1,Ys1) va (Xs2,Ys2)"
    $wsIntel.Range("C2").Value = ""
    $wsIntel.Range("D2").Value = "Xs1"; $wsIntel.Range("E2").Value = "Toa do X moc duoi (Diem 1)"
    $wsIntel.Range("F2").Value = "Ys1"; $wsIntel.Range("G2").Value = "Gia tri Y moc duoi (Diem 1)"
    $wsIntel.Range("H2").Value = "Xs2"; $wsIntel.Range("I2").Value = "Toa do X moc tren (Diem 2)"
    $wsIntel.Range("J2").Value = "Ys2"; $wsIntel.Range("K2").Value = "Gia tri Y moc tren (Diem 2)"
    $wsIntel.Range("L2").Value = "Xs";  $wsIntel.Range("M2").Value = "Gia tri moc X can tra cuu (noi suy)"

    # 2. NoisuyDay
    $wsIntel.Range("A3").Value = "NoisuyDay"
    $wsIntel.Range("B3").Value = "Ham noi suy tuyen tinh day so 1 chieu theo hang hoac cot (tu dong sap xep)"
    $wsIntel.Range("C3").Value = ""
    $wsIntel.Range("D3").Value = "Xk";  $wsIntel.Range("E3").Value = "Day moc tra (Hang hoac cot toa do X)"
    $wsIntel.Range("F3").Value = "Yk";  $wsIntel.Range("G3").Value = "Day gia tri tuong ung (Hang hoac cot gia tri Y)"
    $wsIntel.Range("H3").Value = "Xu";  $wsIntel.Range("I3").Value = "Gia tri moc can tra cuu"

    # 3. NoisuyMatran
    $wsIntel.Range("A4").Value = "NoisuyMatran"
    $wsIntel.Range("B4").Value = "Ham noi suy ma tran 2 chieu song tuyen tinh (Bilinear) bang 2 truc X va Y"
    $wsIntel.Range("C4").Value = ""
    $wsIntel.Range("D4").Value = "Xk";  $wsIntel.Range("E4").Value = "Hang moc truc ngang (X)"
    $wsIntel.Range("F4").Value = "Yk";  $wsIntel.Range("G4").Value = "Cot moc truc doc (Y)"
    $wsIntel.Range("H4").Value = "GT";  $wsIntel.Range("I4").Value = "Vung ma tran so lieu tuong ung"
    $wsIntel.Range("J4").Value = "Xu";  $wsIntel.Range("K4").Value = "Gia tri moc X can tra cuu"
    $wsIntel.Range("L4").Value = "Yu";  $wsIntel.Range("M4").Value = "Gia tri moc Y can tra cuu"

    # 4. Dt
    $wsIntel.Range("A5").Value = "Dt"
    $wsIntel.Range("B5").Value = "Tinh dien tich cot thep tron (mm -> m2): Dt = pi*d^2 / 4000000"
    $wsIntel.Range("C5").Value = ""
    $wsIntel.Range("D5").Value = "d";   $wsIntel.Range("E5").Value = "Duong kinh thanh thep (mm)"

    # An sheet _IntelliSense_
    $wsIntel.Visible = 2
    
    # 1. Add Mod_FontViet_Core
    $modCore = $wb.VBProject.VBComponents.Add(1)
    $modCore.Name = "Mod_FontViet_Core"
    $codeCore = [System.IO.File]::ReadAllText('{bas_core}', [System.Text.Encoding]::UTF8)
    $modCore.CodeModule.AddFromString($codeCore)

    # 2. Add Mod_DValue
    $modDval = $wb.VBProject.VBComponents.Add(1)
    $modDval.Name = "Mod_DValue"
    $codeDval = [System.IO.File]::ReadAllText('{bas_dval}', [System.Text.Encoding]::ASCII)
    $modDval.CodeModule.AddFromString($codeDval)

    # 3. Add Mod_BatchPrint
    $modPrint = $wb.VBProject.VBComponents.Add(1)
    $modPrint.Name = "Mod_BatchPrint"
    $codePrint = [System.IO.File]::ReadAllText('{bas_print}', [System.Text.Encoding]::ASCII)
    $modPrint.CodeModule.AddFromString($codePrint)

    # 3b. Add Mod_NoiSuy
    $modNoiSuy = $wb.VBProject.VBComponents.Add(1)
    $modNoiSuy.Name = "Mod_NoiSuy"
    $codeNoiSuy = [System.IO.File]::ReadAllText('{bas_noisuy}', [System.Text.Encoding]::ASCII)
    $modNoiSuy.CodeModule.AddFromString($codeNoiSuy)
    
    # 4. Add UserForm frmAVTool
    $form1 = $wb.VBProject.VBComponents.Add(3)
    $form1.Name = "frmAVTool"
    $form1.Properties.Item("Caption").Value = "AVTool - Chuyen Doi Font Tieng Viet Pro"
    $form1.Properties.Item("Width").Value = 440
    $form1.Properties.Item("Height").Value = 460

    # Header Title Label
    $lblTitle = $form1.Designer.Controls.Add("Forms.Label.1", "lblHeader")
    $lblTitle.Caption = "AVTOOL - CHUYEN DOI FONT TIENG VIET"
    $lblTitle.Left = 16; $lblTitle.Top = 10; $lblTitle.Width = 400; $lblTitle.Height = 20

    $lblSub = $form1.Designer.Controls.Add("Forms.Label.1", "lblSub")
    $lblSub.Caption = "Tu dong nhan dien TCVN3, VNI -> Unicode | An toan cong thuc 100%"
    $lblSub.Left = 16; $lblSub.Top = 30; $lblSub.Width = 400; $lblSub.Height = 18

    # Frame 1: Bang ma nguon
    $frSrc = $form1.Designer.Controls.Add("Forms.Frame.1", "fraSource")
    $frSrc.Caption = " Bang ma nguon "
    $frSrc.Left = 14; $frSrc.Top = 52; $frSrc.Width = 200; $frSrc.Height = 118

    $opt1 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcAuto")
    $opt1.Caption = "Tu dong nhan dien"
    $opt1.Left = 8; $opt1.Top = 18; $opt1.Width = 185

    $opt2 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcAbc")
    $opt2.Caption = "TCVN3 (.VnTime...)"
    $opt2.Left = 8; $opt2.Top = 42; $opt2.Width = 185

    $opt3 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcVni")
    $opt3.Caption = "VNI-Windows"
    $opt3.Left = 8; $opt3.Top = 66; $opt3.Width = 185

    $opt4 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcUni")
    $opt4.Caption = "Unicode"
    $opt4.Left = 8; $opt4.Top = 90; $opt4.Width = 185

    # Frame 2: Bang ma dich
    $frDst = $form1.Designer.Controls.Add("Forms.Frame.1", "fraTarget")
    $frDst.Caption = " Bang ma dich "
    $frDst.Left = 222; $frDst.Top = 52; $frDst.Width = 200; $frDst.Height = 118

    $optD1 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstUni")
    $optD1.Caption = "Unicode (Chuan)"
    $optD1.Left = 8; $optD1.Top = 22; $optD1.Width = 185

    $optD2 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstAbc")
    $optD2.Caption = "TCVN3 (.VnTime)"
    $optD2.Left = 8; $optD2.Top = 52; $optD2.Width = 185

    $optD3 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstVni")
    $optD3.Caption = "VNI-Windows"
    $optD3.Left = 8; $optD3.Top = 82; $optD3.Width = 185

    # Frame 3: Pham vi
    $frScope = $form1.Designer.Controls.Add("Forms.Frame.1", "fraScope")
    $frScope.Caption = " Pham vi ap dung "
    $frScope.Left = 14; $frScope.Top = 176; $frScope.Width = 408; $frScope.Height = 50

    $opSc1 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeSelection")
    $opSc1.Caption = "Vung dang chon"
    $opSc1.Left = 14; $opSc1.Top = 18; $opSc1.Width = 130

    $opSc2 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeSheet")
    $opSc2.Caption = "Sheet hien tai"
    $opSc2.Left = 155; $opSc2.Top = 18; $opSc2.Width = 125

    $opSc3 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeWorkbook")
    $opSc3.Caption = "Toan bo File"
    $opSc3.Left = 285; $opSc3.Top = 18; $opSc3.Width = 115

    # Frame 4: An Toan
    $frSafe = $form1.Designer.Controls.Add("Forms.Frame.1", "fraSafety")
    $frSafe.Caption = " Tuy chon an toan & Bao ve "
    $frSafe.Left = 14; $frSafe.Top = 232; $frSafe.Width = 408; $frSafe.Height = 65

    $chk1 = $frSafe.Controls.Add("Forms.CheckBox.1", "chkFormula")
    $chk1.Caption = "Bao ve cong thuc (100% khong mat cong thuc)"
    $chk1.Left = 12; $chk1.Top = 16; $chk1.Width = 385

    $chk2 = $frSafe.Controls.Add("Forms.CheckBox.1", "chkShapes")
    $chk2.Caption = "Bo qua hinh anh & do hoa (Giu nguyen anh)"
    $chk2.Left = 12; $chk2.Top = 38; $chk2.Width = 385

    # Frame 5: Kieu chu
    $frCase = $form1.Designer.Controls.Add("Forms.Frame.1", "fraCase")
    $frCase.Caption = " Kieu chu (Tuy chon) "
    $frCase.Left = 14; $frCase.Top = 303; $frCase.Width = 408; $frCase.Height = 46

    $opC0 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseNone")
    $opC0.Caption = "Giu nguyen"
    $opC0.Left = 12; $opC0.Top = 16; $opC0.Width = 85

    $opC1 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseUpper")
    $opC1.Caption = "IN HOA"
    $opC1.Left = 105; $opC1.Top = 16; $opC1.Width = 115

    $opC2 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseLower")
    $opC2.Caption = "in thuong"
    $opC2.Left = 225; $opC2.Top = 16; $opC2.Width = 80

    $opC3 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseTitle")
    $opC3.Caption = "Hoa Dau Tu"
    $opC3.Left = 310; $opC3.Top = 16; $opC3.Width = 90

    # Buttons Form 1
    $btnExec = $form1.Designer.Controls.Add("Forms.CommandButton.1", "btnExecute")
    $btnExec.Caption = "THUC HIEN CHUYEN DOI"
    $btnExec.Left = 40; $btnExec.Top = 362; $btnExec.Width = 230; $btnExec.Height = 34
    $btnExec.Default = $true

    $btnCl = $form1.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $btnCl.Caption = "Dong"
    $btnCl.Left = 290; $btnCl.Top = 362; $btnCl.Width = 110; $btnCl.Height = 34
    $btnCl.Cancel = $true

    $fcode1 = @'
{code_frmAVTool}
'@
    $form1.CodeModule.AddFromString($fcode1)

    # 5. Add Clean UserForm frmBatchPrint
    $form2 = $wb.VBProject.VBComponents.Add(3)
    $form2.Name = "frmBatchPrint"
    $form2.Properties.Item("Caption").Value = "AVTool - In Nhieu Sheet (Fit All Columns & Paper Size)"
    $form2.Properties.Item("Width").Value = 515
    $form2.Properties.Item("Height").Value = 505

    # Header Title
    $lblH2 = $form2.Designer.Controls.Add("Forms.Label.1", "lblHeader")
    $lblH2.Caption = "AVTOOL - IN NHIEU SHEET (BATCH PRINT)"
    $lblH2.Left = 16; $lblH2.Top = 8; $lblH2.Width = 475; $lblH2.Height = 20

    $lblS2 = $form2.Designer.Controls.Add("Forms.Label.1", "lblSub")
    $lblS2.Caption = "Chon sheet in, co vua chieu ngang trang (Fit All Columns) & chon may in"
    $lblS2.Left = 16; $lblS2.Top = 26; $lblS2.Width = 475; $lblS2.Height = 18

    # Frame Col 1: Source
    $frS1 = $form2.Designer.Controls.Add("Forms.Frame.1", "fraSource")
    $frS1.Caption = " Sheet trong Workbook "
    $frS1.Left = 14; $frS1.Top = 46; $frS1.Width = 185; $frS1.Height = 200

    $lstS = $frS1.Controls.Add("Forms.ListBox.1", "lstSource")
    $lstS.Left = 6; $lstS.Top = 16; $lstS.Width = 171; $lstS.Height = 176
    $lstS.MultiSelect = 1

    # Buttons Transfer (Middle)
    $bA1 = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnAddOne")
    $bA1.Caption = ">"
    $bA1.Left = 207; $bA1.Top = 75; $bA1.Width = 34; $bA1.Height = 26

    $bAA = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnAddAll")
    $bAA.Caption = ">>"
    $bAA.Left = 207; $bAA.Top = 107; $bAA.Width = 34; $bAA.Height = 26

    $bR1 = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnRemoveOne")
    $bR1.Caption = "<"
    $bR1.Left = 207; $bR1.Top = 145; $bR1.Width = 34; $bR1.Height = 26

    $bRA = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnRemoveAll")
    $bRA.Caption = "<<"
    $bRA.Left = 207; $bRA.Top = 177; $bRA.Width = 34; $bRA.Height = 26

    # Frame Col 2: Target
    $frT1 = $form2.Designer.Controls.Add("Forms.Frame.1", "fraTarget")
    $frT1.Caption = " Danh sach Sheet can in "
    $frT1.Left = 249; $frT1.Top = 46; $frT1.Width = 185; $frT1.Height = 200

    $lstT = $frT1.Controls.Add("Forms.ListBox.1", "lstTarget")
    $lstT.Left = 6; $lstT.Top = 16; $lstT.Width = 171; $lstT.Height = 176
    $lstT.MultiSelect = 1

    # Buttons Order (Right)
    $bUp = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnMoveUp")
    $bUp.Caption = "Len"
    $bUp.Left = 442; $bUp.Top = 105; $bUp.Width = 48; $bUp.Height = 28

    $bDown = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnMoveDown")
    $bDown.Caption = "Xuong"
    $bDown.Left = 442; $bDown.Top = 145; $bDown.Width = 48; $bDown.Height = 28

    # Options CheckBox All Visible
    $chkVis = $form2.Designer.Controls.Add("Forms.CheckBox.1", "chkAllVisible")
    $chkVis.Caption = "Chi hien thi cac Sheet khong bi an"
    $chkVis.Left = 14; $chkVis.Top = 250; $chkVis.Width = 300; $chkVis.Height = 18
    $chkVis.Value = $true

    # Frame Options: Chon theo mau Tab (Top: 270)
    $lbCol = $form2.Designer.Controls.Add("Forms.Label.1", "lblTabColor")
    $lbCol.Caption = "Chon nhanh theo mau Tab:"
    $lbCol.Left = 14; $lbCol.Top = 272; $lbCol.Width = 150; $lbCol.Height = 18

    $cbCol = $form2.Designer.Controls.Add("Forms.ComboBox.1", "cboTabColor")
    $cbCol.Left = 168; $cbCol.Top = 270; $cbCol.Width = 215; $cbCol.Height = 22
    $cbCol.ColumnCount = 2
    $cbCol.ColumnWidths = "200 pt;0 pt"

    $btCol = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnApplyColor")
    $btCol.Caption = "Chuyen sang in >>"
    $btCol.Left = 388; $btCol.Top = 269; $btCol.Width = 105; $btCol.Height = 24

    # Frame Options: Kho giay, Fit & May in
    $frOpt = $form2.Designer.Controls.Add("Forms.Frame.1", "fraOptions")
    $frOpt.Caption = " Tuy chon in an & May in "
    $frOpt.Left = 14; $frOpt.Top = 298; $frOpt.Width = 478; $frOpt.Height = 112

    # Kho giay
    $lbPa = $frOpt.Controls.Add("Forms.Label.1", "lblPaper")
    $lbPa.Caption = "Kho giay:"
    $lbPa.Left = 14; $lbPa.Top = 20; $lbPa.Width = 70; $lbPa.Height = 18

    $cbPa = $frOpt.Controls.Add("Forms.ComboBox.1", "cboPaperSize")
    $cbPa.Left = 88; $cbPa.Top = 17; $cbPa.Width = 220; $cbPa.Height = 22

    # Fit all columns to 1 page
    $ckFt = $frOpt.Controls.Add("Forms.CheckBox.1", "chkFitAllColumns")
    $ckFt.Caption = "Co vua chieu ngang 1 trang (Fit All Columns on One Page)"
    $ckFt.Left = 14; $ckFt.Top = 48; $ckFt.Width = 450; $ckFt.Height = 22

    # May in
    $lbPrT = $frOpt.Controls.Add("Forms.Label.1", "lblPrinterTitle")
    $lbPrT.Caption = "May in:"
    $lbPrT.Left = 14; $lbPrT.Top = 78; $lbPrT.Width = 55; $lbPrT.Height = 18

    $lbPrN = $frOpt.Controls.Add("Forms.Label.1", "lblPrinterName")
    $lbPrN.Caption = "Dang lay ten may in mac dinh..."
    $lbPrN.Left = 70; $lbPrN.Top = 78; $lbPrN.Width = 290; $lbPrN.Height = 18

    $btPrC = $frOpt.Controls.Add("Forms.CommandButton.1", "btnChangePrinter")
    $btPrC.Caption = "Doi may in..."
    $btPrC.Left = 368; $btPrC.Top = 73; $btPrC.Width = 98; $btPrC.Height = 26

    # Bottom Actions Buttons
    $bPrev = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnPreview")
    $bPrev.Caption = "Xem truoc (Print Preview)"
    $bPrev.Left = 90; $bPrev.Top = 418; $bPrev.Width = 195; $bPrev.Height = 34
    $bPrev.Default = $true

    $bPrn = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnPrint")
    $bPrn.Caption = "In ngay (Print Out)"
    $bPrn.Left = 295; $bPrn.Top = 418; $bPrn.Width = 110; $bPrn.Height = 34

    $bCl2 = $form2.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $bCl2.Caption = "Dong"
    $bCl2.Left = 415; $bCl2.Top = 418; $bCl2.Width = 74; $bCl2.Height = 34
    $bCl2.Cancel = $true

    $fcode2 = @'
{code_frmBatchPrint}
'@
    $form2.CodeModule.AddFromString($fcode2)

    # =========================================================================
    # FORM 3: frmNumberSeparator (Professional Number Separators Switch Form)
    # =========================================================================
    $form3 = $wb.VBProject.VBComponents.Add(3)
    $form3.Name = "frmNumberSeparator"
    $form3.Properties.Item("Width").Value = 440
    $form3.Properties.Item("Height").Value = 265
    $form3.Properties.Item("Caption").Value = "AVTool - Doi Dau Phan Cach So"

    # Header
    $hSep = $form3.Designer.Controls.Add("Forms.Label.1", "lblHeader")
    $hSep.Caption = "DOI DAU PHAN CACH SO"
    $hSep.Left = 16; $hSep.Top = 12; $hSep.Width = 400; $hSep.Height = 18

    $sSep = $form3.Designer.Controls.Add("Forms.Label.1", "lblSub")
    $sSep.Caption = "Chuyen doi dau thap phan & hang nghin tuc thi trong Excel"
    $sSep.Left = 16; $sSep.Top = 32; $sSep.Width = 400; $sSep.Height = 16

    # Frame
    $frSep = $form3.Designer.Controls.Add("Forms.Frame.1", "fraOptions")
    $frSep.Caption = " Tuy chon phan cach "
    $frSep.Left = 16; $frSep.Top = 52; $frSep.Width = 398; $frSep.Height = 135

    # Radio options
    $opVN = $frSep.Controls.Add("Forms.OptionButton.1", "optVN")
    $opVN.Caption = "Chuan Viet Nam: 1.234.567,89 (Phay thap phan, Cham hang nghin)"
    $opVN.Left = 16; $opVN.Top = 18; $opVN.Width = 370; $opVN.Height = 22

    $opUS = $frSep.Controls.Add("Forms.OptionButton.1", "optUS")
    $opUS.Caption = "Chuan Quoc te/US: 1,234,567.89 (Cham thap phan, Phay hang nghin)"
    $opUS.Left = 16; $opUS.Top = 44; $opUS.Width = 370; $opUS.Height = 22

    $opSys = $frSep.Controls.Add("Forms.OptionButton.1", "optSystem")
    $opSys.Caption = "Theo cai dat Windows (Use system separators)"
    $opSys.Left = 16; $opSys.Top = 70; $opSys.Width = 370; $opSys.Height = 22

    $lbCur = $frSep.Controls.Add("Forms.Label.1", "lblCurrent")
    $lbCur.Caption = "Hien tai: ..."
    $lbCur.Left = 16; $lbCur.Top = 98; $lbCur.Width = 370; $lbCur.Height = 28
    $lbCur.ForeColor = 0x808080

    # Buttons
    $btApp = $form3.Designer.Controls.Add("Forms.CommandButton.1", "btnApply")
    $btApp.Caption = "Ap dung"
    $btApp.Left = 210; $btApp.Top = 196; $btApp.Width = 100; $btApp.Height = 28
    $btApp.Default = $true

    $btCls = $form3.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $btCls.Caption = "Dong"
    $btCls.Left = 320; $btCls.Top = 196; $btCls.Width = 94; $btCls.Height = 28
    $btCls.Cancel = $true

    $fcode3 = @'
{code_frmNumberSeparator}
'@
    $form3.CodeModule.AddFromString($fcode3)

    # =========================================================================
    # FORM 4: frmFixDValue (Fix DValue & Formula External Links Form)
    # =========================================================================
    $form4 = $wb.VBProject.VBComponents.Add(3)
    $form4.Name = "frmFixDValue"
    $form4.Properties.Item("Width").Value = 460
    $form4.Properties.Item("Height").Value = 405
    $form4.Properties.Item("Caption").Value = "AVTool - Sua Loi Duong Dan DValue & Cong Thuc"

    # Header
    $hFix = $form4.Designer.Controls.Add("Forms.Label.1", "lblHeader")
    $hFix.Caption = "SUA LOI DUONG DAN DVALUE (#NAME?)"
    $hFix.Left = 16; $hFix.Top = 10; $hFix.Width = 420; $hFix.Height = 18

    $sFix = $form4.Designer.Controls.Add("Forms.Label.1", "lblSub")
    $sFix.Caption = "Tu dong cat bo duong dan file la, khoi phuc cong thuc DValue tuc thi"
    $sFix.Left = 16; $sFix.Top = 28; $sFix.Width = 420; $sFix.Height = 16

    # Frame 1: Status
    $frSta = $form4.Designer.Controls.Add("Forms.Frame.1", "fraStatus")
    $frSta.Caption = " Tinh trang quet trong Workbook "
    $frSta.Left = 16; $frSta.Top = 46; $frSta.Width = 418; $frSta.Height = 102

    $lbSta = $frSta.Controls.Add("Forms.Label.1", "lblStatusSummary")
    $lbSta.Caption = "Dang quet tinh trang..."
    $lbSta.Left = 12; $lbSta.Top = 16; $lbSta.Width = 395; $lbSta.Height = 18

    $txDet = $frSta.Controls.Add("Forms.TextBox.1", "txtDetails")
    $txDet.Left = 12; $txDet.Top = 34; $txDet.Width = 395; $txDet.Height = 58
    $txDet.MultiLine = $true
    $txDet.ScrollBars = 2
    $txDet.Locked = $true
    $txDet.SpecialEffect = 0
    $txDet.BorderStyle = 1

    # Frame 2: Scope (Vertical Layout - completely spacious)
    $frSc = $form4.Designer.Controls.Add("Forms.Frame.1", "fraScope")
    $frSc.Caption = " Pham vi thuc hien "
    $frSc.Left = 16; $frSc.Top = 154; $frSc.Width = 418; $frSc.Height = 86

    $opScS = $frSc.Controls.Add("Forms.OptionButton.1", "optScopeSheet")
    $opScS.Caption = "Sheet hien tai"
    $opScS.Left = 12; $opScS.Top = 16; $opScS.Width = 395; $opScS.Height = 20

    $opScW = $frSc.Controls.Add("Forms.OptionButton.1", "optScopeWorkbook")
    $opScW.Caption = "Toan bo cac Sheet trong File (Workbook)"
    $opScW.Left = 12; $opScW.Top = 38; $opScW.Width = 395; $opScW.Height = 20

    $opScSl = $frSc.Controls.Add("Forms.OptionButton.1", "optScopeSelection")
    $opScSl.Caption = "Chi vung o dang chon (Selection)"
    $opScSl.Left = 12; $opScSl.Top = 60; $opScSl.Width = 395; $opScSl.Height = 20

    # Frame 3: Options
    $frOptD = $form4.Designer.Controls.Add("Forms.Frame.1", "fraOptions")
    $frOptD.Caption = " Tuy chon sua loi "
    $frOptD.Left = 16; $frOptD.Top = 246; $frOptD.Width = 418; $frOptD.Height = 76

    $ckCln = $frOptD.Controls.Add("Forms.CheckBox.1", "chkCleanPath")
    $ckCln.Caption = "Xoa duong dan ngoai dinh vao DValue (Dua ve =DValue())"
    $ckCln.Left = 12; $ckCln.Top = 15; $ckCln.Width = 395; $ckCln.Height = 18

    $ckOth = $frOptD.Controls.Add("Forms.CheckBox.1", "chkNormalizeOthers")
    $ckOth.Caption = "Chuan hoa cac ham khoi luong cu khac (=Tinh, =Eval, =DienGiai)"
    $ckOth.Left = 12; $ckOth.Top = 33; $ckOth.Width = 395; $ckOth.Height = 18

    $ckRec = $frOptD.Controls.Add("Forms.CheckBox.1", "chkRecalculate")
    $ckRec.Caption = "Tu dong tinh toan lai toan bo sau khi sua (CalculateFullRebuild)"
    $ckRec.Left = 12; $ckRec.Top = 51; $ckRec.Width = 395; $ckRec.Height = 18

    # Bottom Buttons
    $btRsc = $form4.Designer.Controls.Add("Forms.CommandButton.1", "btnRescan")
    $btRsc.Caption = "Quet lai"
    $btRsc.Left = 16; $btRsc.Top = 330; $btRsc.Width = 95; $btRsc.Height = 28

    $btFx = $form4.Designer.Controls.Add("Forms.CommandButton.1", "btnFix")
    $btFx.Caption = "SUA LOI NGAY"
    $btFx.Left = 205; $btFx.Top = 330; $btFx.Width = 130; $btFx.Height = 28
    $btFx.Default = $true

    $btClD = $form4.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $btClD.Caption = "Dong"
    $btClD.Left = 344; $btClD.Top = 330; $btClD.Width = 90; $btClD.Height = 28
    $btClD.Cancel = $true

    $fcode4 = @'
{code_frmFixDValue}
'@
    $form4.CodeModule.AddFromString($fcode4)

    # 5. Add UserForm frmNoiSuyMatran
    $form5 = $wb.VBProject.VBComponents.Add(3)
    $form5.Name = "frmNoiSuyMatran"
    $form5.Properties.Item("Caption").Value = "AVTool_NoiSuyMatran_Win"
    $form5.Properties.Item("Width").Value = 460
    $form5.Properties.Item("Height").Value = 435

    $hMat = $form5.Designer.Controls.Add("Forms.Label.1", "lblHeader")
    $hMat.Caption = "NOI SUY MA TRAN 2 CHIEU (BANG TRA 2 TRUC X & Y)"
    $hMat.Left = 16; $hMat.Top = 8; $hMat.Width = 425; $hMat.Height = 18

    $sMat = $form5.Designer.Controls.Add("Forms.Label.1", "lblSub")
    $sMat.Caption = "Tra cuu song tuyen tinh (Bilinear). Tu dong khop kich thuoc & sinh cong thuc"
    $sMat.Left = 16; $sMat.Top = 26; $sMat.Width = 425; $sMat.Height = 16

    $frTb5 = $form5.Designer.Controls.Add("Forms.Frame.1", "fraTable")
    $frTb5.Caption = " Thiet lap vung bang tra ma tran "
    $frTb5.Left = 14; $frTb5.Top = 46; $frTb5.Width = 422; $frTb5.Height = 114

    $lbX5 = $frTb5.Controls.Add("Forms.Label.1", "lblRangeX")
    $lbX5.Caption = "Hang moc X (Ngang):"
    $lbX5.Left = 10; $lbX5.Top = 18; $lbX5.Width = 125; $lbX5.Height = 16
    $txX5 = $frTb5.Controls.Add("Forms.TextBox.1", "txtRangeX")
    $txX5.Left = 138; $txX5.Top = 15; $txX5.Width = 238; $txX5.Height = 20
    $btX5 = $frTb5.Controls.Add("Forms.CommandButton.1", "btnPickX")
    $btX5.Caption = "..."; $btX5.Left = 380; $btX5.Top = 15; $btX5.Width = 32; $btX5.Height = 20

    $lbY5 = $frTb5.Controls.Add("Forms.Label.1", "lblRangeY")
    $lbY5.Caption = "Cot moc Y (Doc):"
    $lbY5.Left = 10; $lbY5.Top = 42; $lbY5.Width = 125; $lbY5.Height = 16
    $txY5 = $frTb5.Controls.Add("Forms.TextBox.1", "txtRangeY")
    $txY5.Left = 138; $txY5.Top = 39; $txY5.Width = 238; $txY5.Height = 20
    $btY5 = $frTb5.Controls.Add("Forms.CommandButton.1", "btnPickY")
    $btY5.Caption = "..."; $btY5.Left = 380; $btY5.Top = 39; $btY5.Width = 32; $btY5.Height = 20

    $lbGT5 = $frTb5.Controls.Add("Forms.Label.1", "lblRangeGT")
    $lbGT5.Caption = "Ma tran so lieu (GT):"
    $lbGT5.Left = 10; $lbGT5.Top = 66; $lbGT5.Width = 125; $lbGT5.Height = 16
    $txGT5 = $frTb5.Controls.Add("Forms.TextBox.1", "txtRangeGT")
    $txGT5.Left = 138; $txGT5.Top = 63; $txGT5.Width = 238; $txGT5.Height = 20
    $btGT5 = $frTb5.Controls.Add("Forms.CommandButton.1", "btnPickGT")
    $btGT5.Caption = "..."; $btGT5.Left = 380; $btGT5.Top = 63; $btGT5.Width = 32; $btGT5.Height = 20

    $lbDim5 = $frTb5.Controls.Add("Forms.Label.1", "lblDimCheck")
    $lbDim5.Caption = "Chua chon vung du lieu"
    $lbDim5.Left = 138; $lbDim5.Top = 88; $lbDim5.Width = 270; $lbDim5.Height = 18

    $frQ5 = $form5.Designer.Controls.Add("Forms.Frame.1", "fraQuery")
    $frQ5.Caption = " Gia tri can tra cuu & O dat ket qua "
    $frQ5.Left = 14; $frQ5.Top = 164; $frQ5.Width = 422; $frQ5.Height = 72

    $lbVx5 = $frQ5.Controls.Add("Forms.Label.1", "lblValX")
    $lbVx5.Caption = "Moc Xu:"
    $lbVx5.Left = 10; $lbVx5.Top = 16; $lbVx5.Width = 48; $lbVx5.Height = 16
    $txVx5 = $frQ5.Controls.Add("Forms.TextBox.1", "txtValX")
    $txVx5.Left = 60; $txVx5.Top = 14; $txVx5.Width = 85; $txVx5.Height = 20
    $btVx5 = $frQ5.Controls.Add("Forms.CommandButton.1", "btnPickValX")
    $btVx5.Caption = "..."; $btVx5.Left = 148; $btVx5.Top = 14; $btVx5.Width = 26; $btVx5.Height = 20

    $lbVy5 = $frQ5.Controls.Add("Forms.Label.1", "lblValY")
    $lbVy5.Caption = "Moc Yu:"
    $lbVy5.Left = 185; $lbVy5.Top = 16; $lbVy5.Width = 48; $lbVy5.Height = 16
    $txVy5 = $frQ5.Controls.Add("Forms.TextBox.1", "txtValY")
    $txVy5.Left = 235; $txVy5.Top = 14; $txVy5.Width = 85; $txVy5.Height = 20
    $btVy5 = $frQ5.Controls.Add("Forms.CommandButton.1", "btnPickValY")
    $btVy5.Caption = "..."; $btVy5.Left = 323; $btVy5.Top = 14; $btVy5.Width = 26; $btVy5.Height = 20

    $lbDst5 = $frQ5.Controls.Add("Forms.Label.1", "lblDest")
    $lbDst5.Caption = "O ket qua:"
    $lbDst5.Left = 10; $lbDst5.Top = 42; $lbDst5.Width = 70; $lbDst5.Height = 16
    $txDst5 = $frQ5.Controls.Add("Forms.TextBox.1", "txtDest")
    $txDst5.Left = 85; $txDst5.Top = 40; $txDst5.Width = 160; $txDst5.Height = 20
    $btDst5 = $frQ5.Controls.Add("Forms.CommandButton.1", "btnPickDest")
    $btDst5.Caption = "..."; $btDst5.Left = 248; $btDst5.Top = 40; $btDst5.Width = 28; $btDst5.Height = 20

    $ckAbs5 = $frQ5.Controls.Add("Forms.CheckBox.1", "chkAbsolute")
    $ckAbs5.Caption = "Co dinh bang ($)"
    $ckAbs5.Left = 285; $ckAbs5.Top = 41; $ckAbs5.Width = 125; $ckAbs5.Height = 18
    $ckAbs5.Value = $true

    $frPv5 = $form5.Designer.Controls.Add("Forms.Frame.1", "fraPreview")
    $frPv5.Caption = " Xem truoc ket qua & Cong thuc sinh ra "
    $frPv5.Left = 14; $frPv5.Top = 240; $frPv5.Width = 422; $frPv5.Height = 85

    $lbPv5 = $frPv5.Controls.Add("Forms.Label.1", "lblPreview")
    $lbPv5.Caption = "Ket qua xem truoc: (Cho nhap thong tin...)"
    $lbPv5.Left = 10; $lbPv5.Top = 14; $lbPv5.Width = 400; $lbPv5.Height = 16

    $txFm5 = $frPv5.Controls.Add("Forms.TextBox.1", "txtFormula")
    $txFm5.Left = 10; $txFm5.Top = 32; $txFm5.Width = 402; $txFm5.Height = 44
    $txFm5.MultiLine = $true; $txFm5.Locked = $true; $txFm5.BorderStyle = 1

    $btSmp5 = $form5.Designer.Controls.Add("Forms.CommandButton.1", "btnSample")
    $btSmp5.Caption = "Tao Bang Mau"
    $btSmp5.Left = 14; $btSmp5.Top = 336; $btSmp5.Width = 115; $btSmp5.Height = 30

    $btIns5 = $form5.Designer.Controls.Add("Forms.CommandButton.1", "btnInsert")
    $btIns5.Caption = "CHEN CONG THUC"
    $btIns5.Left = 138; $btIns5.Top = 336; $btIns5.Width = 185; $btIns5.Height = 30
    $btIns5.Default = $true

    $btCls5 = $form5.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $btCls5.Caption = "Dong"
    $btCls5.Left = 332; $btCls5.Top = 336; $btCls5.Width = 104; $btCls5.Height = 30
    $btCls5.Cancel = $true

    $fcode5 = @'
{code_frmNoiSuyMatran}
'@
    $form5.CodeModule.AddFromString($fcode5)

    # 6. Add UserForm frmNoiSuyDay
    $form6 = $wb.VBProject.VBComponents.Add(3)
    $form6.Name = "frmNoiSuyDay"
    $form6.Properties.Item("Caption").Value = "AVTool_NoiSuyDay_Win"
    $form6.Properties.Item("Width").Value = 460
    $form6.Properties.Item("Height").Value = 400

    $hDay = $form6.Designer.Controls.Add("Forms.Label.1", "lblHeader")
    $hDay.Caption = "NOI SUY DAY SO 1 CHIEU (TUYEN TINH)"
    $hDay.Left = 16; $hDay.Top = 8; $hDay.Width = 425; $hDay.Height = 18

    $sDay = $form6.Designer.Controls.Add("Forms.Label.1", "lblSub")
    $sDay.Caption = "Tra cuu tuyen tinh theo day hang hoac cot. Tu dong sap xep moc tang dan"
    $sDay.Left = 16; $sDay.Top = 26; $sDay.Width = 425; $sDay.Height = 16

    $frTb6 = $form6.Designer.Controls.Add("Forms.Frame.1", "fraTable")
    $frTb6.Caption = " Thiet lap 2 day du lieu bang tra "
    $frTb6.Left = 14; $frTb6.Top = 46; $frTb6.Width = 422; $frTb6.Height = 90

    $lbX6 = $frTb6.Controls.Add("Forms.Label.1", "lblRangeX")
    $lbX6.Caption = "Day moc tra (Day X):"
    $lbX6.Left = 10; $lbX6.Top = 18; $lbX6.Width = 125; $lbX6.Height = 16
    $txX6 = $frTb6.Controls.Add("Forms.TextBox.1", "txtRangeX")
    $txX6.Left = 138; $txX6.Top = 15; $txX6.Width = 238; $txX6.Height = 20
    $btX6 = $frTb6.Controls.Add("Forms.CommandButton.1", "btnPickX")
    $btX6.Caption = "..."; $btX6.Left = 380; $btX6.Top = 15; $btX6.Width = 32; $btX6.Height = 20

    $lbY6 = $frTb6.Controls.Add("Forms.Label.1", "lblRangeY")
    $lbY6.Caption = "Day gia tri (Day Y):"
    $lbY6.Left = 10; $lbY6.Top = 42; $lbY6.Width = 125; $lbY6.Height = 16
    $txY6 = $frTb6.Controls.Add("Forms.TextBox.1", "txtRangeY")
    $txY6.Left = 138; $txY6.Top = 39; $txY6.Width = 238; $txY6.Height = 20
    $btY6 = $frTb6.Controls.Add("Forms.CommandButton.1", "btnPickY")
    $btY6.Caption = "..."; $btY6.Left = 380; $btY6.Top = 39; $btY6.Width = 32; $btY6.Height = 20

    $lbDim6 = $frTb6.Controls.Add("Forms.Label.1", "lblDimCheck")
    $lbDim6.Caption = "Chua chon day du lieu"
    $lbDim6.Left = 138; $lbDim6.Top = 64; $lbDim6.Width = 270; $lbDim6.Height = 18

    $frQ6 = $form6.Designer.Controls.Add("Forms.Frame.1", "fraQuery")
    $frQ6.Caption = " Gia tri can tra cuu & O dat ket qua "
    $frQ6.Left = 14; $frQ6.Top = 140; $frQ6.Width = 422; $frQ6.Height = 72

    $lbVx6 = $frQ6.Controls.Add("Forms.Label.1", "lblValX")
    $lbVx6.Caption = "Gia tri Xu:"
    $lbVx6.Left = 10; $lbVx6.Top = 16; $lbVx6.Width = 70; $lbVx6.Height = 16
    $txVx6 = $frQ6.Controls.Add("Forms.TextBox.1", "txtValX")
    $txVx6.Left = 85; $txVx6.Top = 14; $txVx6.Width = 160; $txVx6.Height = 20
    $btVx6 = $frQ6.Controls.Add("Forms.CommandButton.1", "btnPickValX")
    $btVx6.Caption = "..."; $btVx6.Left = 248; $btVx6.Top = 14; $btVx6.Width = 28; $btVx6.Height = 20

    $lbDst6 = $frQ6.Controls.Add("Forms.Label.1", "lblDest")
    $lbDst6.Caption = "O ket qua:"
    $lbDst6.Left = 10; $lbDst6.Top = 42; $lbDst6.Width = 70; $lbDst6.Height = 16
    $txDst6 = $frQ6.Controls.Add("Forms.TextBox.1", "txtDest")
    $txDst6.Left = 85; $txDst6.Top = 40; $txDst6.Width = 160; $txDst6.Height = 20
    $btDst6 = $frQ6.Controls.Add("Forms.CommandButton.1", "btnPickDest")
    $btDst6.Caption = "..."; $btDst6.Left = 248; $btDst6.Top = 40; $btDst6.Width = 28; $btDst6.Height = 20

    $ckAbs6 = $frQ6.Controls.Add("Forms.CheckBox.1", "chkAbsolute")
    $ckAbs6.Caption = "Co dinh day ($)"
    $ckAbs6.Left = 285; $ckAbs6.Top = 41; $ckAbs6.Width = 125; $ckAbs6.Height = 18
    $ckAbs6.Value = $true

    $frPv6 = $form6.Designer.Controls.Add("Forms.Frame.1", "fraPreview")
    $frPv6.Caption = " Xem truoc ket qua & Cong thuc sinh ra "
    $frPv6.Left = 14; $frPv6.Top = 216; $frPv6.Width = 422; $frPv6.Height = 85

    $lbPv6 = $frPv6.Controls.Add("Forms.Label.1", "lblPreview")
    $lbPv6.Caption = "Ket qua xem truoc: (Cho nhap thong tin...)"
    $lbPv6.Left = 10; $lbPv6.Top = 14; $lbPv6.Width = 400; $lbPv6.Height = 16

    $txFm6 = $frPv6.Controls.Add("Forms.TextBox.1", "txtFormula")
    $txFm6.Left = 10; $txFm6.Top = 32; $txFm6.Width = 402; $txFm6.Height = 44
    $txFm6.MultiLine = $true; $txFm6.Locked = $true; $txFm6.BorderStyle = 1

    $btSmp6 = $form6.Designer.Controls.Add("Forms.CommandButton.1", "btnSample")
    $btSmp6.Caption = "Tao Day Mau"
    $btSmp6.Left = 14; $btSmp6.Top = 312; $btSmp6.Width = 115; $btSmp6.Height = 30

    $btIns6 = $form6.Designer.Controls.Add("Forms.CommandButton.1", "btnInsert")
    $btIns6.Caption = "CHEN CONG THUC"
    $btIns6.Left = 138; $btIns6.Top = 312; $btIns6.Width = 185; $btIns6.Height = 30
    $btIns6.Default = $true

    $btCls6 = $form6.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $btCls6.Caption = "Dong"
    $btCls6.Left = 332; $btCls6.Top = 312; $btCls6.Width = 104; $btCls6.Height = 30
    $btCls6.Cancel = $true

    $fcode6 = @'
{code_frmNoiSuyDay}
'@
    $form6.CodeModule.AddFromString($fcode6)

    if (Test-Path '{temp_xlam}') {{ Remove-Item '{temp_xlam}' }}
    $wb.SaveAs('{temp_xlam}', 55)
    $wb.Close($false)
    Write-Host "Temp XLAM generated with Form & DValue & BatchPrint Clean."
}} catch {{
    Write-Host "Error: " $_.Exception.Message
    exit 1
}} finally {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""

with open(os.path.join(workspace, "make_master.ps1"), "w", encoding="utf-8-sig") as f:
    f.write(ps_script)

res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", os.path.join(workspace, "make_master.ps1")], capture_output=True, text=True)
print("PowerShell Output:", res.stdout)
if res.returncode != 0:
    print("PowerShell Error:", res.stderr)
    sys.exit(1)

# Clean Ribbon XML: imageMso="FilePrintQuick"
ribbon_xml_2007 = """<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabAVTool" label="AVTool">
        <group id="grpAVTool" label="Đổi Font Tiếng Việt">
          <button id="btnOpenForm" label="Đổi Font Tiếng Việt" size="large" imageMso="FontDialog" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpDValue" label="Dự Toán Khối Lượng">
          <button id="btnInsertDValue" label="Chèn DValue Vào File" size="large" imageMso="Calculator" onAction="Ribbon_OnAction"/>
          <button id="btnFixDValue" label="Sửa Lỗi DValue (#NAME?)" size="large" imageMso="FormulaEvaluate" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpNoiSuy" label="Nội Suy Bảng Tra">
          <button id="btnNoiSuyMatran" label="Nội Suy Ma Trận 2D" size="large" imageMso="PivotTableInsert" onAction="Ribbon_OnAction"/>
          <button id="btnNoiSuyDay" label="Nội Suy Dãy 1D" size="large" imageMso="AutoSum" onAction="Ribbon_OnAction"/>
          <button id="btnNoiSuy1Chieu" label="Nội Suy 1 Chiều" size="large" imageMso="EquationProfessional" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpSheetUtils" label="Tiện Ích Sheet &amp; Số">
          <button id="btnChangeSeparators" label="Đổi Dấu Số (, / .)" size="large" imageMso="CommaStyle" onAction="Ribbon_OnAction"/>
          <button id="btnUnhideRows" label="Mở Dòng Ẩn Bôi Màu" size="large" imageMso="OutlineShowDetail" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpPrint" label="In Ấn">
          <button id="btnBatchPrint" label="In Nhiều Sheet" size="large" imageMso="FilePrintQuick" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpUpdate" label="Hệ Thống">
          <button id="btnCheckUpdate" label="Cập Nhật Bản Mới" size="large" imageMso="ServerRefresh" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

ribbon_xml_2010 = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabAVTool" label="AVTool">
        <group id="grpAVTool" label="Đổi Font Tiếng Việt">
          <button id="btnOpenForm" label="Đổi Font Tiếng Việt" size="large" imageMso="FontDialog" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpDValue" label="Dự Toán Khối Lượng">
          <button id="btnInsertDValue" label="Chèn DValue Vào File" size="large" imageMso="Calculator" onAction="Ribbon_OnAction"/>
          <button id="btnFixDValue" label="Sửa Lỗi DValue (#NAME?)" size="large" imageMso="FormulaEvaluate" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpNoiSuy" label="Nội Suy Bảng Tra">
          <button id="btnNoiSuyMatran" label="Nội Suy Ma Trận 2D" size="large" imageMso="PivotTableInsert" onAction="Ribbon_OnAction"/>
          <button id="btnNoiSuyDay" label="Nội Suy Dãy 1D" size="large" imageMso="AutoSum" onAction="Ribbon_OnAction"/>
          <button id="btnNoiSuy1Chieu" label="Nội Suy 1 Chiều" size="large" imageMso="EquationProfessional" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpSheetUtils" label="Tiện Ích Sheet &amp; Số">
          <button id="btnChangeSeparators" label="Đổi Dấu Số (, / .)" size="large" imageMso="CommaStyle" onAction="Ribbon_OnAction"/>
          <button id="btnUnhideRows" label="Mở Dòng Ẩn Bôi Màu" size="large" imageMso="OutlineShowDetail" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpPrint" label="In Ấn">
          <button id="btnBatchPrint" label="In Nhiều Sheet" size="large" imageMso="FilePrintQuick" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpUpdate" label="Hệ Thống">
          <button id="btnCheckUpdate" label="Cập Nhật Bản Mới" size="large" imageMso="ServerRefresh" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

for attempt in range(15):
    subprocess.run(["powershell", "-Command", "Stop-Process -Name 'EXCEL' -Force -ErrorAction SilentlyContinue"])
    time.sleep(1)
    try:
        with open(temp_xlam, 'r+b') as check_f:
            break
    except IOError:
        pass

files_data = {}
with zipfile.ZipFile(temp_xlam, 'r') as zin:
    for item in zin.infolist():
        files_data[item.filename] = zin.read(item.filename)

files_data['customUI/customUI.xml'] = ribbon_xml_2007.encode('utf-8')
files_data['customUI/customUI14.xml'] = ribbon_xml_2010.encode('utf-8')

rels_content = files_data['_rels/.rels'].decode('utf-8')
rel1 = '<Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2006/01/relationships/ui/extensibility" Target="customUI/customUI.xml"/>'
rel2 = '<Relationship Id="rIdCustomUI14" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'

if 'customUI.xml' not in rels_content:
    rels_content = rels_content.replace('</Relationships>', f'  {rel1}\n  {rel2}\n</Relationships>')
    files_data['_rels/.rels'] = rels_content.encode('utf-8')

with zipfile.ZipFile(final_xlam, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for fname, fdata in files_data.items():
        zout.writestr(fname, fdata)

print(f"Successfully generated master {final_xlam}!")

# Update AddIns folder
addin_dir = os.path.expandvars(r"%APPDATA%\Microsoft\AddIns")
dest = os.path.join(addin_dir, "TN_FontViet_Pro_2026.xlam")
shutil.copy2(final_xlam, dest)
print("Updated AddIns location:", dest)

updater_bat = os.path.join(workspace, "Cap_Nhat_Online.bat")
if os.path.exists(updater_bat):
    shutil.copy2(updater_bat, os.path.join(addin_dir, "Cap_Nhat_Online.bat"))
    print("Updated Cap_Nhat_Online.bat in AddIns folder.")

# Clean temp files
for tf in [temp_xlam, os.path.join(workspace, "make_master.ps1")]:
    if os.path.exists(tf):
        os.remove(tf)
