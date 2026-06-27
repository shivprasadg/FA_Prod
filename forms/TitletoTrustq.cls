Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ckoutschbtn_Click()
    Dim MissingX As String
    MissingX = "| "
    If IsNull(Me.SelClient) Then MissingX = MissingX & "Customer | "
        If IsNull(SelLease) Then MissingX = MissingX & "Lease | "
          If IsNull(SelSch) Then MissingX = MissingX & "Schedule | "
        If IsNull(SentDate) Then MissingX = MissingX & "Date Sent | "
    If MissingX <> "| " Then
        MsgBox "You MUST complete All the fields on the form" & vbNewLine & "You're missing " & MissingX, vbCritical, "Missing Info"
        Exit Sub
    End If
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "TitleSentToTrust"
    DoCmd.SetWarnings True
End Sub
