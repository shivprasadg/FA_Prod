Attribute VB_Name = "SchSumRptTLSPDF"
Option Compare Database


Sub OpenSchSumRptTLSPDF()

 'Removed LBJ 5/30/2019 *********************************************
    'Dim strFileName As String
    '  Dim i As String

    'j = Format(Now, "YYYYDD HMS")

    'Const acFormatPDF = "PDF Format (*.pdf)"
    'DoCmd.SetWarnings False

    'strFileName = "Schedule Summary TLS" & j & ".pdf"

    'DoCmd.OpenReport "SchSummaryRptTLS", acViewPreview
    'DoCmd.OutputTo acOutputReport, "SchSummaryRptTLS", acFormatPDF, strFileName, True
    'DoCmd.close acReport, "SchSummaryRptTLS", acSaveNo

    'DoCmd.SetWarnings True

End Sub

Function FileInUse(strFileName As String) As Boolean
   On Error Resume Next
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   If Err.Number <> 0 Then
      MsgBox ("Report is already open, please close and rerun"), vbOKOnly
      FileInUse = True
      Err.Clear
   End If
End Function
