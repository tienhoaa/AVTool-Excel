import os, sys, zipfile, subprocess, shutil

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
bas_file = os.path.join(workspace, "Mod_FontViet_Core.bas")
temp_xlsm = os.path.join(workspace, "temp_xlsm.xlsm")
final_xlsm = os.path.join(workspace, "AVTool_FontViet.xlsm")

form_code = '''
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    optSrcAuto.Value = True
    optDstUni.Value = True
    optScopeSelection.Value = True
    chkFormula.Value = True
    chkShapes.Value = True
    optCaseNone.Value = True
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
    
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "HuongDan"
    $ws.Range("B2").Value2 = "CHAO MUNG DEN VOI AVTOOL - CHUYEN DOI FONT TIENG VIET"
    $ws.Range("B2").Font.Bold = $true
    $ws.Range("B2").Font.Size = 14
    $ws.Range("B4").Value2 = "Hay nhin len thanh Ribbon tren cung, bam vao Tab [AVTool] de su dung!"
    $ws.Range("B5").Value2 = "Bam nut [Chuyen Doi Font (Mo Form)] de bat hop thoai dieu khien."

    # 1. Add Code Module
    $module = $wb.VBProject.VBComponents.Add(1)
    $module.Name = "Mod_FontViet_Core"
    $code = [System.IO.File]::ReadAllText('{bas_file}', [System.Text.Encoding]::UTF8)
    $module.CodeModule.AddFromString($code)
    
    # 2. Add UserForm
    $form = $wb.VBProject.VBComponents.Add(3)
    $form.Name = "frmAVTool"
    $form.Properties.Item("Caption").Value = "AVTool - Chuyen Doi Font Tieng Viet"
    $form.Properties.Item("Width").Value = 370
    $form.Properties.Item("Height").Value = 405

    # Frame 1: Source
    $frSrc = $form.Designer.Controls.Add("Forms.Frame.1", "fraSource")
    $frSrc.Caption = " Bang ma nguon "
    $frSrc.Left = 10; $frSrc.Top = 10; $frSrc.Width = 165; $frSrc.Height = 110

    $opt1 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcAuto")
    $opt1.Caption = "Tu dong nhan dien"
    $opt1.Left = 8; $opt1.Top = 15; $opt1.Width = 145

    $opt2 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcAbc")
    $opt2.Caption = "TCVN3 (.VnTime...)"
    $opt2.Left = 8; $opt2.Top = 38; $opt2.Width = 145

    $opt3 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcVni")
    $opt3.Caption = "VNI-Windows"
    $opt3.Left = 8; $opt3.Top = 61; $opt3.Width = 145

    $opt4 = $frSrc.Controls.Add("Forms.OptionButton.1", "optSrcUni")
    $opt4.Caption = "Unicode"
    $opt4.Left = 8; $opt4.Top = 84; $opt4.Width = 145

    # Frame 2: Target
    $frDst = $form.Designer.Controls.Add("Forms.Frame.1", "fraTarget")
    $frDst.Caption = " Bang ma dich "
    $frDst.Left = 185; $frDst.Top = 10; $frDst.Width = 165; $frDst.Height = 110

    $optD1 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstUni")
    $optD1.Caption = "Unicode (Chuan)"
    $optD1.Left = 8; $optD1.Top = 20; $optD1.Width = 145

    $optD2 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstAbc")
    $optD2.Caption = "TCVN3 (.VnTime)"
    $optD2.Left = 8; $optD2.Top = 50; $optD2.Width = 145

    $optD3 = $frDst.Controls.Add("Forms.OptionButton.1", "optDstVni")
    $optD3.Caption = "VNI-Windows"
    $optD3.Left = 8; $optD3.Top = 80; $optD3.Width = 145

    # Frame 3: Scope
    $frScope = $form.Designer.Controls.Add("Forms.Frame.1", "fraScope")
    $frScope.Caption = " Pham vi ap dung "
    $frScope.Left = 10; $frScope.Top = 125; $frScope.Width = 340; $frScope.Height = 50

    $opSc1 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeSelection")
    $opSc1.Caption = "Vung chon"
    $opSc1.Left = 10; $opSc1.Top = 18; $opSc1.Width = 90

    $opSc2 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeSheet")
    $opSc2.Caption = "Sheet hien tai"
    $opSc2.Left = 110; $opSc2.Top = 18; $opSc2.Width = 105

    $opSc3 = $frScope.Controls.Add("Forms.OptionButton.1", "optScopeWorkbook")
    $opSc3.Caption = "Toan bo File"
    $opSc3.Left = 230; $opSc3.Top = 18; $opSc3.Width = 95

    # Frame 4: Safety
    $frSafe = $form.Designer.Controls.Add("Forms.Frame.1", "fraSafety")
    $frSafe.Caption = " Tuy chon an toan "
    $frSafe.Left = 10; $frSafe.Top = 180; $frSafe.Width = 340; $frSafe.Height = 65

    $chk1 = $frSafe.Controls.Add("Forms.CheckBox.1", "chkFormula")
    $chk1.Caption = "Bao ve cong thuc (100% khong mat cong thuc)"
    $chk1.Left = 10; $chk1.Top = 16; $chk1.Width = 310

    $chk2 = $frSafe.Controls.Add("Forms.CheckBox.1", "chkShapes")
    $chk2.Caption = "Bo qua hinh anh va Shapes (Khong anh huong anh)"
    $chk2.Left = 10; $chk2.Top = 38; $chk2.Width = 310

    # Frame 5: Case
    $frCase = $form.Designer.Controls.Add("Forms.Frame.1", "fraCase")
    $frCase.Caption = " Kieu chu (Tuy chon) "
    $frCase.Left = 10; $frCase.Top = 250; $frCase.Width = 340; $frCase.Height = 45

    $opC0 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseNone")
    $opC0.Caption = "Giu nguyen"
    $opC0.Left = 10; $opC0.Top = 16; $opC0.Width = 80

    $opC1 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseUpper")
    $opC1.Caption = "IN HOA"
    $opC1.Left = 95; $opC1.Top = 16; $opC1.Width = 75

    $opC2 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseLower")
    $opC2.Caption = "in thuong"
    $opC2.Left = 175; $opC2.Top = 16; $opC2.Width = 75

    $opC3 = $frCase.Controls.Add("Forms.OptionButton.1", "optCaseTitle")
    $opC3.Caption = "Hoa Dau Tu"
    $opC3.Left = 255; $opC3.Top = 16; $opC3.Width = 80

    # Buttons
    $btnExec = $form.Designer.Controls.Add("Forms.CommandButton.1", "btnExecute")
    $btnExec.Caption = "THUC HIEN CHUYEN DOI"
    $btnExec.Left = 30; $btnExec.Top = 310; $btnExec.Width = 190; $btnExec.Height = 32
    $btnExec.Default = $true

    $btnCl = $form.Designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    $btnCl.Caption = "Dong"
    $btnCl.Left = 240; $btnCl.Top = 310; $btnCl.Width = 90; $btnCl.Height = 32
    $btnCl.Cancel = $true

    $fcode = @'
{form_code}
'@
    $form.CodeModule.AddFromString($fcode)

    if (Test-Path '{temp_xlsm}') {{ Remove-Item '{temp_xlsm}' }}
    $wb.SaveAs('{temp_xlsm}', 52) # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    $wb.Close($false)
    Write-Host "Temp XLSM generated."
}} catch {{
    Write-Host "Error: " $_.Exception.Message
    exit 1
}} finally {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""

with open(os.path.join(workspace, "make_xlsm.ps1"), "w", encoding="utf-8") as f:
    f.write(ps_script)

res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", os.path.join(workspace, "make_xlsm.ps1")], capture_output=True, text=True)
print("PowerShell Output:", res.stdout)
if res.returncode != 0:
    print("PowerShell Error:", res.stderr)
    sys.exit(1)

ribbon_xml_2007 = """<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabAVTool" label="AVTool">
        <group id="grpForm" label="Bang Dieu Khien">
          <button id="btnOpenForm" label="Chuyen Doi Font (Mo Form)" size="large" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpScope" label="1. Pham Vi">
          <button id="btnScopeSelection" label="Vung Dang Chon" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnScopeSheet" label="Sheet Hien Tai" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnScopeWorkbook" label="Toan Bo File" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpToUnicode" label="2. Sang Unicode">
          <button id="btnAutoToUni" label="Tu Dong Nhan Dien" size="large" onAction="Ribbon_OnAction"/>
          <button id="btnAbcToUni" label="TCVN3 (.Vn) -> Unicode" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnVniToUni" label="VNI-Win -> Unicode" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpFromUnicode" label="3. Tu Unicode">
          <button id="btnUniToAbc" label="Unicode -> TCVN3" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnUniToVni" label="Unicode -> VNI" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpSafety" label="4. An Toan &amp; Bao Ve">
          <button id="btnToggleFormula" label="Bao Ve Cong Thuc [BAT]" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnToggleShapes" label="Bo Qua Hinh Anh [BAT]" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpCase" label="5. Doi Kieu Chu">
          <button id="btnUpper" label="CHU HOA" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnLower" label="chu thuong" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnTitle" label="Hoa Dau Tu" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpAbout" label="6. Tro Giup">
          <button id="btnHelp" label="Huong Dan" size="large" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

ribbon_xml_2010 = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabAVTool" label="AVTool">
        <group id="grpForm" label="Bang Dieu Khien">
          <button id="btnOpenForm" label="Chuyen Doi Font (Mo Form)" size="large" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpScope" label="1. Pham Vi">
          <button id="btnScopeSelection" label="Vung Dang Chon" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnScopeSheet" label="Sheet Hien Tai" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnScopeWorkbook" label="Toan Bo File" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpToUnicode" label="2. Sang Unicode">
          <button id="btnAutoToUni" label="Tu Dong Nhan Dien" size="large" onAction="Ribbon_OnAction"/>
          <button id="btnAbcToUni" label="TCVN3 (.Vn) -> Unicode" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnVniToUni" label="VNI-Win -> Unicode" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpFromUnicode" label="3. Tu Unicode">
          <button id="btnUniToAbc" label="Unicode -> TCVN3" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnUniToVni" label="Unicode -> VNI" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpSafety" label="4. An Toan &amp; Bao Ve">
          <button id="btnToggleFormula" label="Bao Ve Cong Thuc [BAT]" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnToggleShapes" label="Bo Qua Hinh Anh [BAT]" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpCase" label="5. Doi Kieu Chu">
          <button id="btnUpper" label="CHU HOA" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnLower" label="chu thuong" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnTitle" label="Hoa Dau Tu" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpAbout" label="6. Tro Giup">
          <button id="btnHelp" label="Huong Dan" size="large" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

files_data = {}
with zipfile.ZipFile(temp_xlsm, 'r') as zin:
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

with zipfile.ZipFile(final_xlsm, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for fname, fdata in files_data.items():
        zout.writestr(fname, fdata)

print(f"Successfully generated {final_xlsm}!")
