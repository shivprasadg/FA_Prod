Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Sub SetValueAdd()
    Me.Form.RecordSource = "SELECT * FROM Concession Where ClientId=" & Me.Parent.ClientID & " Order By ConDt Desc;"
End Sub
Private Sub Form_Close()
   On Error GoTo Form_Close_Error
   On Error GoTo 0
   Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_CONCESS"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

Select Case DataErr
Case 3146 'Required Field ODBC Error 'Case 3146 To 3621 'ODBC Error
MsgBox "You must have data in all Yellow Required Fields before moving off the record or closing the form.", vbCritical, "Required Field Error #" + str(DataErr) + " - All Required Fields Must Contain Data!"
Response = acDataErrContinue
Case 3146 To 3621
'MsgBox "Unknown ODBC error occurred", vbCritical, "ODBC Error #" & Str(DataErr)
'Response = acDataErrContinue
Case Else
'MsgBox AccessError(DataErr)
'Response = acDataErrContinue
End Select
End Sub

Private Sub ValueAddedReport_Click()
DoCmd.OpenReport "ValueAddedReport", acViewPreview
End Sub
