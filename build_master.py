import os, sys, zipfile, subprocess, shutil

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
bas_core = os.path.join(workspace, "Mod_FontViet_Core.bas")
bas_dval = os.path.join(workspace, "Mod_DValue.bas")
temp_xlam = os.path.join(workspace, "temp_addon.xlam")
final_xlam = os.path.join(workspace, "TN_FontViet_Pro_2026.xlam")

subprocess.run(["powershell", "-Command", "Stop-Process -Name 'EXCEL' -Force -ErrorAction SilentlyContinue"])

def esc(text):
    out = []
    for ch in text:
        if ord(ch) > 127:
            out.append(f"\\u{ord(ch):04X}")
        else:
            out.append(ch)
    return "".join(out)

elements = [
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

lines = []
for var, val in elements:
    if var == "Me.Caption":
        lines.append(f'    {var} = "{val}"')
    else:
        lines.append(f'    {var} = UniText("{esc(val)}")')
captions_code = "\n".join(lines)

form_code = f'''
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
    
{captions_code}
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
        MsgBox UniText("Vui l\u00F2ng ch\u1ECDn b\u1EA3ng m\u00E3 c\u1EA7n chuy\u1EC3n \u0111\u1ED5i ho\u1EB7c ki\u1EC3u ch\u1EEF!"), vbExclamation, "AVTool"
        Exit Sub
    End If
    
    Me.Hide
    ExecuteConversion convMode, False, caseType
    Me.Show vbModeless
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
    
    # 3. Add UserForm
    $form = $wb.VBProject.VBComponents.Add(3)
    $form.Name = "frmAVTool"
    $form.Properties.Item("Caption").Value = "AVTool - Chuyen Doi Font Tieng Viet Pro"
    $form.Properties.Item("Width").Value = 440
    $form.Properties.Item("Height").Value = 460

    # Header Title Label
    $lblTitle = $form.Designer.Controls.Add("Forms.Label.1", "lblHeader")
    $lblTitle.Caption = "AVTOOL - CHUYEN DOI FONT TIENG VIET"
    $lblTitle.Left = 16; $lblTitle.Top = 10; $lblTitle.Width = 400; $lblTitle.Height = 20

    $lblSub = $form.Designer.Controls.Add("Forms.Label.1", "lblSub")
    $lblSub.Caption = "Tu dong nhan dien TCVN3, VNI -> Unicode | An toan cong thuc 100%"
    $lblSub.Left = 16; $lblSub.Top = 30; $lblSub.Width = 400; $lblSub.Height = 18

    # Frame 1: Bang ma nguon
    $frSrc = $form.Designer.Controls.Add("Forms.Frame.1", "fraSource")
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
    $frDst = $form.Designer.Controls.Add("Forms.Frame.1", "fraTarget")
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
    $frScope = $form.Designer.Controls.Add("Forms.Frame.1", "fraScope")
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
    $frSafe = $form.Designer.Controls.Add("Forms.Frame.1", "fraSafety")
    $frSafe.Caption = " Tuy chon an toan & Bao ve "
    $frSafe.Left = 14; $frSafe.Top = 232; $frSafe.Width = 408; $frSafe.Height = 65

    $chk1 = $frSafe.Controls.Add("Forms.CheckBox.1", "chkFormula")
    $chk1.Caption = "Bao ve cong thuc (100% khong mat cong thuc)"
    $chk1.Left = 12; $chk1.Top = 16; $chk1.Width = 385

    $chk2 = $frSafe.Controls.Add("Forms.CheckBox.1", "chkShapes")
    $chk2.Caption = "Bo qua hinh anh & do hoa (Giu nguyen anh)"
    $chk2.Left = 12; $chk2.Top = 38; $chk2.Width = 385

    # Frame 5: Kieu chu
    $frCase = $form.Designer.Controls.Add("Forms.Frame.1", "fraCase")
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

    # Buttons
    $btnExec = $form.Designer.Controls.Add("Forms.CommandButton.1", "btnExecute")
    $btnExec.Caption = "THUC HIEN CHUYEN DOI"
    $btnExec.Left = 40; $btnExec.Top = 362; $btnExec.Width = 230; $btnExec.Height = 34
    $btnExec.Default = $true

    $btnCl = $form.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $btnCl.Caption = "Dong"
    $btnCl.Left = 290; $btnCl.Top = 362; $btnCl.Width = 110; $btnCl.Height = 34
    $btnCl.Cancel = $true

    $fcode = @'
{form_code}
'@
    $form.CodeModule.AddFromString($fcode)

    if (Test-Path '{temp_xlam}') {{ Remove-Item '{temp_xlam}' }}
    $wb.SaveAs('{temp_xlam}', 55)
    $wb.Close($false)
    Write-Host "Temp XLAM generated with Form & DValue."
}} catch {{
    Write-Host "Error: " $_.Exception.Message
    exit 1
}} finally {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""

with open(os.path.join(workspace, "make_master.ps1"), "w", encoding="utf-8") as f:
    f.write(ps_script)

res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", os.path.join(workspace, "make_master.ps1")], capture_output=True, text=True)
print("PowerShell Output:", res.stdout)
if res.returncode != 0:
    print("PowerShell Error:", res.stderr)
    sys.exit(1)

# Clean Ribbon XML: 2 groups (Font Chu + Tinh Khoi Luong DValue)
ribbon_xml_2007 = """<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabAVTool" label="AVTool">
        <group id="grpAVTool" label="Doi Font Tieng Viet">
          <button id="btnOpenForm" label="Doi Font Tieng Viet" size="large" imageMso="FontDialog" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpDValue" label="Du Toan &amp; Khoi Luong">
          <button id="btnInsertDValue" label="Chen DValue Vao File" size="large" imageMso="Calculator" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpUpdate" label="He Thong">
          <button id="btnCheckUpdate" label="Cap Nhat Ban Moi" size="large" imageMso="FileUpdate" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

ribbon_xml_2010 = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabAVTool" label="AVTool">
        <group id="grpAVTool" label="Doi Font Tieng Viet">
          <button id="btnOpenForm" label="Doi Font Tieng Viet" size="large" imageMso="FontDialog" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpDValue" label="Du Toan &amp; Khoi Luong">
          <button id="btnInsertDValue" label="Chen DValue Vao File" size="large" imageMso="Calculator" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpUpdate" label="He Thong">
          <button id="btnCheckUpdate" label="Cap Nhat Ban Moi" size="large" imageMso="FileUpdate" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

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

# Clean temp files
for tf in [temp_xlam, os.path.join(workspace, "make_master.ps1"), os.path.join(workspace, "check_dval_icon.py"), os.path.join(workspace, "check_dval_chars.py"), os.path.join(workspace, "convert_dvalue.py"), os.path.join(workspace, "prep_core.py"), os.path.join(workspace, "build_spacious.py")]:
    if os.path.exists(tf):
        os.remove(tf)
