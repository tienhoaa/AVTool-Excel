import os, sys, zipfile, subprocess, shutil, time

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
bas_core = os.path.join(workspace, "Mod_FontViet_Core.bas")
bas_dval = os.path.join(workspace, "Mod_DValue.bas")
bas_print = os.path.join(workspace, "Mod_BatchPrint.bas")
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
    Dim convMode As Integer
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
    
    If optSrcAuto.Value Then
        convMode = 3
    ElseIf optSrcAbc.Value Then
        convMode = 1
    ElseIf optSrcVni.Value Then
        convMode = 2
    ElseIf optSrcUni.Value Then
        If optDstAbc.Value Then
            convMode = 4
        ElseIf optDstVni.Value Then
            convMode = 5
        Else
            convMode = 0
        End If
    End If
    
    If optCaseUpper.Value Then
        caseType = 1
    ElseIf optCaseLower.Value Then
        caseType = 2
    ElseIf optCaseTitle.Value Then
        caseType = 3
    Else
        caseType = 0
    End If
    
    If convMode = 0 And caseType = 0 Then
        MsgBoxUni "{esc('Vui lòng chọn bảng mã cần chuyển đổi hoặc kiểu chữ!')}", vbExclamation, "AVTool"
        Exit Sub
    End If
    
    Me.Hide
    ExecuteConversion convMode, False, caseType
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

ps_script = f"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $wb = $excel.Workbooks.Add()
    
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
        <group id="grpDValue" label="Dự Toán &amp; Khối Lượng">
          <button id="btnInsertDValue" label="Chèn DValue Vào File" size="large" imageMso="Calculator" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpPrint" label="In Ấn">
          <button id="btnBatchPrint" label="In Nhiều Sheet" size="large" imageMso="FilePrintQuick" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpUpdate" label="Hệ Thống">
          <button id="btnCheckUpdate" label="Cập Nhật Bản Mới" size="large" imageMso="FileUpdate" onAction="Ribbon_OnAction"/>
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
        <group id="grpDValue" label="Dự Toán &amp; Khối Lượng">
          <button id="btnInsertDValue" label="Chèn DValue Vào File" size="large" imageMso="Calculator" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpPrint" label="In Ấn">
          <button id="btnBatchPrint" label="In Nhiều Sheet" size="large" imageMso="FilePrintQuick" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpUpdate" label="Hệ Thống">
          <button id="btnCheckUpdate" label="Cập Nhật Bản Mới" size="large" imageMso="FileUpdate" onAction="Ribbon_OnAction"/>
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
