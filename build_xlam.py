import os, sys, zipfile, shutil, subprocess

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
bas_file = os.path.join(workspace, "Mod_FontViet_Core.bas")
temp_xlam = os.path.join(workspace, "temp_addon.xlam")
final_xlam = os.path.join(workspace, "TN_FontViet_Pro_2026.xlam")

subprocess.run(["powershell", "-Command", "Stop-Process -Name 'EXCEL' -Force -ErrorAction SilentlyContinue"])

ps_script = f"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $wb = $excel.Workbooks.Add()
    $module = $wb.VBProject.VBComponents.Add(1)
    $module.Name = "Mod_FontViet_Core"
    $code = [System.IO.File]::ReadAllText('{bas_file}', [System.Text.Encoding]::UTF8)
    $module.CodeModule.AddFromString($code)
    
    if (Test-Path '{temp_xlam}') {{ Remove-Item '{temp_xlam}' }}
    $wb.SaveAs('{temp_xlam}', 55) # 55 = xlOpenXMLAddIn
    $wb.Close($false)
    Write-Host "Temp XLAM generated successfully."
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

# Ribbon XML for Office 2007 (customUI.xml)
ribbon_xml_2007 = """<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabFontViet" label="ĐỔI FONT TIẾNG VIỆT">
        <group id="grpScope" label="1. Phạm Vi">
          <button id="btnScopeSelection" label="Vùng Đang Chọn" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnScopeSheet" label="Sheet Hiện Tại" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnScopeWorkbook" label="Toàn Bộ File" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpToUnicode" label="2. Sang Unicode">
          <button id="btnAutoToUni" label="Tự Động Nhận Diện" size="large" onAction="Ribbon_OnAction"/>
          <separator id="sep1"/>
          <button id="btnAbcToUni" label="TCVN3 (.Vn) -> Unicode" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnVniToUni" label="VNI-Win -> Unicode" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpFromUnicode" label="3. Từ Unicode">
          <button id="btnUniToAbc" label="Unicode -> TCVN3" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnUniToVni" label="Unicode -> VNI" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpSafety" label="4. An Toàn &amp; Bảo Vệ">
          <button id="btnToggleFormula" label="Bảo Vệ Công Thức [BẬT]" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnToggleShapes" label="Bỏ Qua Hình Ảnh [BẬT]" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpCase" label="5. Đổi Kiểu Chữ">
          <button id="btnUpper" label="CHỮ HOA" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnLower" label="chữ thường" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnTitle" label="Hoa Đầu Từ" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpAbout" label="6. Trợ Giúp">
          <button id="btnHelp" label="Hướng Dẫn" size="large" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

# Ribbon XML for Office 2010+ (customUI14.xml)
ribbon_xml_2010 = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabFontViet" label="ĐỔI FONT TIẾNG VIỆT">
        <group id="grpScope" label="1. Phạm Vi">
          <button id="btnScopeSelection" label="Vùng Đang Chọn" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnScopeSheet" label="Sheet Hiện Tại" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnScopeWorkbook" label="Toàn Bộ File" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpToUnicode" label="2. Sang Unicode">
          <button id="btnAutoToUni" label="Tự Động Nhận Diện" size="large" onAction="Ribbon_OnAction"/>
          <separator id="sep1"/>
          <button id="btnAbcToUni" label="TCVN3 (.Vn) -> Unicode" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnVniToUni" label="VNI-Win -> Unicode" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpFromUnicode" label="3. Từ Unicode">
          <button id="btnUniToAbc" label="Unicode -> TCVN3" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnUniToVni" label="Unicode -> VNI" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpSafety" label="4. An Toàn &amp; Bảo Vệ">
          <button id="btnToggleFormula" label="Bảo Vệ Công Thức [BẬT]" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnToggleShapes" label="Bỏ Qua Hình Ảnh [BẬT]" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpCase" label="5. Đổi Kiểu Chữ">
          <button id="btnUpper" label="CHỮ HOA" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnLower" label="chữ thường" size="normal" onAction="Ribbon_OnAction"/>
          <button id="btnTitle" label="Hoa Đầu Từ" size="normal" onAction="Ribbon_OnAction"/>
        </group>
        <group id="grpAbout" label="6. Trợ Giúp">
          <button id="btnHelp" label="Hướng Dẫn" size="large" onAction="Ribbon_OnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

files_data = {}
with zipfile.ZipFile(temp_xlam, 'r') as zin:
    for item in zin.infolist():
        files_data[item.filename] = zin.read(item.filename)

# Add customUI
files_data['customUI/customUI.xml'] = ribbon_xml_2007.encode('utf-8')
files_data['customUI/customUI14.xml'] = ribbon_xml_2010.encode('utf-8')

# Update _rels/.rels
rels_content = files_data['_rels/.rels'].decode('utf-8')
rel1 = '<Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2006/01/relationships/ui/extensibility" Target="customUI/customUI.xml"/>'
rel2 = '<Relationship Id="rIdCustomUI14" Type="http://schemas.microsoft.com/office/2007/06/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'

if 'customUI.xml' not in rels_content:
    rels_content = rels_content.replace('</Relationships>', f'  {rel1}\n  {rel2}\n</Relationships>')
    files_data['_rels/.rels'] = rels_content.encode('utf-8')

# Write final xlam
with zipfile.ZipFile(final_xlam, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for fname, fdata in files_data.items():
        zout.writestr(fname, fdata)

print(f"Successfully generated {final_xlam}!")
