Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RST        As DAO.Recordset
Dim SQLz       As String

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
    Case 3146    'Required Field ODBC Error 'Case 3146 To 3621 'ODBC Error
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

Private Sub ADDML_Click()
    DoCmd.GotoRecord , "", acNewRec
End Sub

Public Sub SetLease()
    Me.Form.RecordSource = ""
    SQLz = "SELECT MLID, Client, MLNo, MLOrig FROM MstrLease WHERE Client=" & Me.Parent.clientID & " ORDER BY MstrLease.MLOrig, MstrLease.MLNo;"
    Me.MasterLeaseList.RowSource = SQLz
    Me.MLAmendf.Form.AmendmentsList.RowSource = ""
End Sub
Private Sub MasterLeaseList_Click()
    SQLz = "SELECT MstrLease.*, Clients.clientTaxID FROM MstrLease LEFT JOIN " & _
           "Clients ON MstrLease.MLGuarantor = Clients.clientID WHERE MLID = " & Me.MasterLeaseList & ";"
    Me.Form.RecordSource = SQLz
    SQLz = "SELECT FullName, Title, contactID, ClientID FROM ClientContacts_Base " & _
           "WHERE ClientID= " & Me.clientID & " Order by FullName;"
    Me.MLSigner.RowSource = SQLz
    Me.MLSigner2.RowSource = SQLz
    Me.MLSigner3.RowSource = SQLz
    SQLz = "SELECT FullName, Title, contactID, ClientID FROM ClientContacts_Base " & _
           "WHERE ClientID= " & Me.MLGuarantor & " Order by FullName;"
    Me.MLGuarSigner.RowSource = SQLz
    SQLz = "SELECT * From mlamends WHERE MLRef= " & MLID & " ORDER BY MLAmdDt DESC;"
    Me.MLAmendf.Form.RecordSource = SQLz
    SQLz = "SELECT ID, MLAmtNO AS AmendNum, MLAmdDt AS AmendDate, MLRef From MLAamends" & _
           "WHERE MLRef =" & Me.MLID & " ORDER BY MLAmtNO DESC , MLAmdDt DESC;"
    Me.MLAmendf.Form.AmendmentsList.RowSource = SQLz
End Sub

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    If IsOpen("ATF_SFImporter") Then
        Forms!ATF_SFImporter!MLID.Requery
        Forms!ATF_SFImporter!MLID = Me.MLID
        Forms!ATF_SFImporter!MLID2.Requery
        Forms!ATF_SFImporter!MLID2 = Me.MLID
    End If
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_Leasing_sub"
End Sub

Private Sub MLGuarantor_Click()
    Me.MLGuarantor.Requery
End Sub

Private Sub MLGuarSigner_Click()
    Me.MLGuarSignerID = Me.MLGuarSigner.Column(2)
End Sub

Private Sub MLLessor_Click()
    Me.MLLessorID = Me.MLLessor.Column(1)
End Sub

Private Sub MLSigner_Click()
    Me.MLSignerID = Me.MLSigner.Column(2)
End Sub
Private Sub MLSigner2_Click()
    Me.MLSigner2ID = Me.MLSigner2.Column(2)
End Sub
Private Sub MLSigner3_Click()
    Me.MLSigner3ID = Me.MLSigner3.Column(2)
End Sub
Private Sub LeaseGuarantor_Change()
    Forms!ClientsView.LeaseGuarantor.Requery
End Sub
