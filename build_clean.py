import os, sys, zipfile, subprocess, shutil

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
bas_file = os.path.join(workspace, "Mod_FontViet_Core.bas")
temp_xlam = os.path.join(workspace, "temp_addon.xlam")
final_xlam = os.path.join(workspace, "TN_FontViet_Pro_2026.xlam")

subprocess.run(["powershell", "-Command", "Stop-Process -Name 'EXCEL' -Force -ErrorAction SilentlyContinue"])

form_code = '''
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
    lblHeader.Font.Size = 10.5
    lblHeader.Font.Bold = True
    lblSub.Font.Size = 8
    btnExecute.Font.Bold = True
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
        MsgBox "Vui long chon bang ma can chuyen doi hoac kieu chu!", vbExclamation, "AVTool"
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
    
    # 1. Add Code Module
    $module = $wb.VBProject.VBComponents.Add(1)
    $module.Name = "Mod_FontViet_Core"
    $code = [System.IO.File]::ReadAllText('{bas_file}', [System.Text.Encoding]::UTF8)
    $module.CodeModule.AddFromString($code)
    
    # 2. Add UserForm
    $form = $wb.VBProject.VBComponents.Add(3)
    $form.Name = "frmAVTool"
    $form.Properties.Item("Caption").Value = "AVTool - Chuyen Doi Font Tieng Viet Pro"
    $form.Properties.Item("Width").Value = 385
    $form.Properties.Item("Height").Value = 445

    # Header Title Label
    $lblTitle = $form.Designer.Controls.Add("Forms.Label.1", "lblHeader")
    $lblTitle.Caption = "AVTOOL - CHUYEN DOI FONT TIENG VIET"
    $lblTitle.Left = 14; $lblTitle.Top = 10; $lblTitle.Width = 350; $lblTitle.Height = 18

    $lblSub = $form.Designer.Controls.Add("Forms.Label.1", "lblSub")
    $lblSub.Caption = "Tu dong nhan dien TCVN3, VNI -> Unicode | An toan cong thuc 100%"
    $lblSub.Left = 14; $lblSub.Top = 28; $lblSub.Width = 350; $lblSub.Height = 16

    # Frame 1: Bang ma nguon
    $frSrc = $form.Designer.Controls.Add("Forms.Frame.1", "fraSource")
    $frSrc.Caption = " Bang ma nguon "
    $frSrc.Left = 12; $frSrc.Top = 48; $frSrc.Width = 172; $frSrc.Height = 115

    $opt1 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcAuto")
    $opt1.Caption = "Tu dong nhan dien"
    $opt1.Left = 8; $opt1.Top = 18; $opt1.Width = 152

    $opt2 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcAbc")
    $opt2.Caption = "TCVN3 (.VnTime...)"
    $opt2.Left = 8; $opt2.Top = 42; $opt2.Width = 152

    $opt3 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcVni")
    $opt3.Caption = "VNI-Windows"
    $opt3.Left = 8; $opt3.Top = 66; $opt3.Width = 152

    $opt4 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcUni")
    $opt4.Caption = "Unicode"
    $opt4.Left = 8; $opt4.Top = 90; $opt4.Width = 152

    # Frame 2: Bang ma dich
    $frDst = $form.Designer.Controls.Add("Forms.Frame.1", "fraTarget")
    $frDst.Caption = " Bang ma dich "
    $frDst.Left = 192; $frDst.Top = 48; $frDst.Width = 172; $frDst.Height = 115

    $optD1 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstUni")
    $optD1.Caption = "Unicode (Chuan)"
    $optD1.Left = 8; $optD1.Top = 22; $optD1.Width = 152

    $optD2 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstAbc")
    $optD2.Caption = "TCVN3 (.VnTime)"
    $optD2.Left = 8; $optD2.Top = 52; $optD2.Width = 152

    $optD3 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstVni")
    $optD3.Caption = "VNI-Windows"
    $optD3.Left = 8; $optD3.Top = 82; $optD3.Width = 152

    # Frame 3: Pham vi
    $frScope = $form.Designer.Controls.Add("Forms.Frame.1", "fraScope")
    $frScope.Caption = " Pham vi ap dung "
    $frScope.Left = 12; $frScope.Top = 168; $frScope.Width = 352; $frScope.Height = 50

    $opSc1 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeSelection")
    $opSc1.Caption = "Vung dang chon"
    $opSc1.Left = 10; $opSc1.Top = 18; $opSc1.Width = 110

    $opSc2 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeSheet")
    $opSc2.Caption = "Sheet hien tai"
    $opSc2.Left = 125; $opSc2.Top = 18; $opSc2.Width = 105

    $opSc3 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeWorkbook")
    $opSc3.Caption = "Toan bo File"
    $opSc3.Left = 240; $opSc3.Top = 18; $opSc3.Width = 100

    # Frame 4: An Toan
    $frSafe = $form.Designer.Controls.Add("Forms.Frame.1", "fraSafety")
    $frSafe.Caption = " Tuy chon an toan & Bao ve "
    $frSafe.Left = 12; $frSafe.Top = 224; $frSafe.Width = 352; $frSafe.Height = 65

    $chk1 = $frSafe.Controls.Add("Forms.CheckBox.1", "chkFormula")
    $chk1.Caption = "Bao ve cong thuc (100% khong mat cong thuc)"
    $chk1.Left = 10; $chk1.Top = 16; $chk1.Width = 330

    $chk2 = $frSafe.Controls.Add("Forms.CheckBox.1", "chkShapes")
    $chk2.Caption = "Bo qua hinh anh & do hoa (Giu nguyen anh)"
    $chk2.Left = 10; $chk2.Top = 38; $chk2.Width = 330

    # Frame 5: Kieu chu
    $frCase = $form.Designer.Controls.Add("Forms.Frame.1", "fraCase")
    $frCase.Caption = " Kieu chu (Tuy chon) "
    $frCase.Left = 12; $frCase.Top = 295; $frCase.Width = 352; $frCase.Height = 45

    $opC0 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseNone")
    $opC0.Caption = "Giu nguyen"
    $opC0.Left = 10; $opC0.Top = 16; $opC0.Width = 85

    $opC1 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseUpper")
    $opC1.Caption = "IN HOA"
    $opC1.Left = 100; $opC1.Top = 16; $opC1.Width = 75

    $opC2 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseLower")
    $opC2.Caption = "in thuong"
    $opC2.Left = 180; $opC2.Top = 16; $opC2.Width = 75

    $opC3 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseTitle")
    $opC3.Caption = "Hoa Dau Tu"
    $opC3.Left = 260; $opC3.Top = 16; $opC3.Width = 80

    # Buttons
    $btnExec = $form.Designer.Controls.Add("Forms.CommandButton.1", "btnExecute")
    $btnExec.Caption = "THUC HIEN CHUYEN DOI"
    $btnExec.Left = 25; $btnExec.Top = 352; $btnExec.Width = 205; $btnExec.Height = 32
    $btnExec.Default = $true

    $btnCl = $form.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $btnCl.Caption = "Dong"
    $btnCl.Left = 245; $btnCl.Top = 352; $btnCl.Width = 95; $btnCl.Height = 32
    $btnCl.Cancel = $true

    $fcode = @'
{form_code}
'@
    $form.CodeModule.AddFromString($fcode)

    if (Test-Path '{temp_xlam}') {{ Remove-Item '{temp_xlam}' }}
    $wb.SaveAs('{temp_xlam}', 55)
    $wb.Close($false)
    Write-Host "Temp XLAM generated with Form."
}} catch {{
    Write-Host "Error: " $_.Exception.Message
    exit 1
}} finally {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""

with open(os.path.join(workspace, "make_temp.ps1"), "w", encoding="utf-8") as f:
    f.write(ps_script)

res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", os.path.join(workspace, "make_temp.ps1")], capture_output=True, text=True)
print("PowerShell Output:", res.stdout)
if res.returncode != 0:
    print("PowerShell Error:", res.stderr)
    sys.exit(1)

# Clean Ribbon XML: ONLY 1 main button in AVTool tab
ribbon_xml_2007 = """<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabAVTool" label="AVTool">
        <group id="grpAVTool" label="Cong Cu Font">
          <button id="btnOpenForm" label="Doi Font Tieng Viet" size="large" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

ribbon_xml_2010 = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabAVTool" label="AVTool">
        <group id="grpAVTool" label="Cong Cu Font">
          <button id="btnOpenForm" label="Doi Font Tieng Viet" size="large" onAction="Ribbon_OnAction"/>
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

print(f"Successfully generated {final_xlam}!")

# Update AddIns folder
addin_dir = os.path.expandvars(r"%APPDATA%\Microsoft\AddIns")
dest = os.path.join(addin_dir, "TN_FontViet_Pro_2026.xlam")
shutil.copy2(final_xlam, dest)
print("Updated AddIns location:", dest)

# Clean temp files
for tf in [temp_xlam, os.path.join(workspace, "make_temp.ps1"), os.path.join(workspace, "test_font.ps1")]:
    if os.path.exists(tf):
        os.remove(tf)
