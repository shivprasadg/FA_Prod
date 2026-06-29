Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz As String

Private Sub CloseMe_Click()
If CheckAcceptedDate = True Then
DoCmd.OpenForm "FAMainMenu"
If AllUnitsInGroup_Accepted(Me.SchID) = True Then AllUnitsInGroup_AcceptedMessage (Me.SchID)
DoCmd.close acForm, Me.Form.Name
Else
MsgBox "You Cannot Close The Form", vbCritical, "Missing Data"
End If
End Sub

Private Function CheckAcceptedDate() As Boolean
    Dim RST    As DAO.Recordset
    Dim x      As Integer
    CheckAcceptedDate = True
    SQLz = Me.Unit_List_CofASub.Form.RecordSource
    SQLz = Replace(SQLz, "ORDER", " Where sgrpid=" & Me.SchID & " ORDER", 1, 1)
       Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveFirst
    For x = 1 To RST.RecordCount
        If IsDate(Nz(RST![UnitCOABack], RST!UnitCofACopy)) And Not IsDate(RST!UnitAcceptDate) Then
            x = RST.RecordCount
            MsgBox "You Must enter an Accepted Date!" & vbNewLine & "VIN:" & RST!unitvin, vbCritical, "Missing Data"
            CheckAcceptedDate = False
        Else
            RST.MoveNext
        End If
    Next x
End Function
Private Sub Form_Load()
Me.HeaderLabel.Visible = False
Me.InstructionLabel2.Visible = False
Me.InstructionLabel1.Visible = True
Me.Unit_List_CofASub.Visible = False
End Sub


Private Sub MasterLease_Click()
SQLz = "SELECT SGrpID, SchID, " & "[Schedule] & ' | ' & [UnitGroup]" & " AS [Schedule|Group] " & _
       "From vw_SixKeys " & _
       "WHERE [MLID] =" & MasterLease.Column(0) & " " & _
       "GROUP BY SGrpID, SchID, Schedule, UnitGroup " & _
       "ORDER BY SchID DESC, SGrpID ;"

'SQLz = "SELECT SchID, SGrpID, 'Sch: '" & Schedule & "' - Grp: '" & UnitGroup & "' FROM vw_SixKeys WHERE MLID=" & MasterLease.Column(0) & " Group BY SchID, SGrpID ORDER BY SGrpID DESC;"
Me.SchID.RowSource = SQLz
Me.SchID.Selected(0) = True
Me.Unit_List_CofASub.Visible = False
Me.InstructionLabel2.Visible = True
Me.InstructionLabel1.Visible = False
End Sub

Private Sub SchID_Click()
Me.SearchLease = ""
Me.Unit_List_CofASub.Visible = True
Me.HeaderLabel.Visible = True
End Sub

Private Sub SearchLease_Change()
SQLz = "SELECT MLID, MLNo FROM vw_SixKeys " & _
"WHERE MLOrig='FA' AND (status='A' Or status='B') AND MLNo Like '*" & Me.SearchLease.Text & "*' " & _
"GROUP BY MLID, MLNo ORDER BY MLNo;"

'SQLz = "SELECT MLID, MLNo " & _
       "From MstrLease " & _
       "WHERE MLNo Not Like 'demo*' AND MLOrig='fa' AND MLNo Like '*" & Me.SearchLease.Text & "*'" & _
       "ORDER BY MstrLease.MLNo;"
Me.MasterLease.RowSource = SQLz
End Sub
