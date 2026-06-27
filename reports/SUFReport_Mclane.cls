Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub ExportPDF_Click()
DoCmd.OutputTo acOutputReport, “”, acFormatPDF, , False
End Sub
Private Sub Report_Open(Cancel As Integer)
    If Not IsNull(Me.OpenArgs) Then Me.Report.RecordSource = Me.OpenArgs
End Sub
Private Sub Report_Load()
    If Nz(Me.clientID, "") = 4971 Then
        Me.McLaneGrey.Height = 2555
        Me.McLaneLine.Visible = True
        Me.McLaneBox.Visible = True
        coName = "McLane"
    Else
        Me.McLaneGrey.Height = 2248
        Me.McLaneLine.Visible = False
        Me.McLaneBox.Visible = False
        coName = Me.CompanyName
    End If
End Sub
