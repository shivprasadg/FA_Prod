Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub ExportPDF_Click()
DoCmd.OutputTo acOutputReport, “”, acFormatPDF, , False
End Sub


Private Sub Report_Load()
Exit Sub
'If IsNull(Me.RefferInst) Then
'Me.ReeferBox.Height = 0
'Me.ReeferBox.Visible = False
'Else
'Me.ReeferBox.Height = (1440 * 0.531)
'Me.ReeferBox.Visible = True
'End If
End Sub

Private Sub Report_Open(Cancel As Integer)
  If Not IsNull(Me.OpenArgs) Then Me.Report.RecordSource = Me.OpenArgs
End Sub
