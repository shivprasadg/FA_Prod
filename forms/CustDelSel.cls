Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command10_Click()

Dim objExcel As Object

Set objExcel = CreateObject("Excel.Application")
On Error Resume Next
    If (Eval("[Forms]![custdelsel]![selgrp] Is Not Null And [Forms]![custdelsel]![selcust]")) Then
        Beep
        MsgBox "Select Group or Specific Customer not both", vbCritical, "Selection Error"
        Exit Sub
    End If

    If (Forms!CustDelSel!selprov = "aim" Or Forms!CustDelSel!selprov = "penske") Then
        DoCmd.OpenQuery "DelRptQMainesprov", acViewNormal, acEdit
        DoCmd.OutputTo acOutputQuery, "DelRptQMainesprov", "Excel97-Excel2003Workbook(*.xls)", "", True, "", 0, acExportQualityPrint
objExcel.Selection.Columns.AutoFit

    End If
    If (Eval("[Forms]![custdelsel]![selprov] Is Null")) Then
        DoCmd.OpenQuery "DelRptQMainesA", acViewNormal, acEdit
        DoCmd.OutputTo acOutputQuery, "DelRptQMainesA", "Excel97-Excel2003Workbook(*.xls)", "", True, "", 0, acExportQualityPrint
objExcel.Selection.Columns.AutoFit


    End If

End Sub
