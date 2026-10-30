Attribute VB_Name = "Mod_BatchPrint"
Option Explicit

Public Sub ShowBatchPrintForm()
    On Error Resume Next
    If ActiveWorkbook Is Nothing Then
        MsgBox "Vui long mo mot file Excel truoc khi su dung chuc nang nay!", vbExclamation, "AVTool - In Nhieu Sheet"
        Exit Sub
    End If
    
    If ActiveWorkbook.Sheets.Count = 0 Then
        MsgBox "File Excel hien tai khong co sheet nao!", vbExclamation, "AVTool - In Nhieu Sheet"
        Exit Sub
    End If
    
    frmBatchPrint.Show vbModeless
End Sub
