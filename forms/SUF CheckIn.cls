Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Public SQLz As String

Private Sub CloseMe_Click()
DoCmd.OpenForm "FAMainMenu"
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
Me.Unit_List_SUFSub.Visible = False
End Sub



Private Sub MasterLease_AfterUpdate()
SQLz = "SELECT Schedule.SchID, Schedule.SchNo " & _
"FROM MstrLease INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef " & _
"WHERE MstrLease.MLID = " & Me.MasterLease & " " & _
"ORDER BY Schedule.SchID DESC;"
Me.SchID.RowSource = SQLz
Me.SchID.Selected(0) = True
Me.Unit_List_SUFSub.Visible = False


End Sub

Private Sub SchID_Click()
Me.SearchLease = ""
Me.Unit_List_SUFSub.Visible = True
End Sub

Private Sub SearchLease_Change()
SQLz = "SELECT MLID, MLNo " & _
       "From MstrLease " & _
       "WHERE MLNo Not Like 'demo*' AND MLNo Like '*" & Me.SearchLease.Text & "*'" & _
       "ORDER BY MstrLease.MLNo;"
Me.MasterLease.RowSource = SQLz
End Sub
