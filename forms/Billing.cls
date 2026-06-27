Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub GroupList_Click()
Me.groupID = Me.GroupList.Column(0)
Me.BillingSubForm.Form.Filter = "GroupID = " & Me.GroupList.Column(0) & ""
Me.BillingSubForm.Form.FilterOn = True
End Sub

Private Sub RunAutoExtend__Click()
  Call AutoExtendMTM(False)
End Sub

Private Sub UpdateBilling_Click()
If Me.SchID & "xyz" = "xyz" Then Exit Sub
DoCmd.OpenForm "Schedule_Form", acNormal, , "SchId=" & Me.SchID & ""
End Sub

Private Sub Form_Load()
  If DLookup("[Access Level]", "TBL_Employees", "[User ID] = FAUserName()") = 5 Then
     '  Me.BillInfoBtn.Visible = False
       Me.[10bill].Visible = False
       Me.[18bill].Visible = False
      Else
     '   Me.BillInfoBtn.Visible = True
       End If
Me.SelLease.RowSource = ""
Me.clientID = Null
Me.MLNo = Null
Me.SchID = Null
Me.SchNo = Null
Me.groupID = Null

End Sub

Private Sub SelClient_Click()
SQLz = "SELECT MLID, MLNo, Client From MstrLease " & _
       "WHERE MLOrig = 'fa' and Client=" & Me.SelClient.Column(0) & " ORDER BY MstrLease.MLNo;"
Me.SelLease.RowSource = SQLz
Me.clientID = Me.SelClient.Column(0)
Me.SelSch.RowSource = ""
Me.MLNo = Null
Me.SchID = Null
Me.SchNo = Null
Me.groupID = Null
Me.GroupList.Requery
End Sub

Private Sub SelLease_Click()
SQLz = "SELECT S.SchID, S.SchNo, iif(G.Strp >0,""Y"","""") as Strp, G.mBLCD as BLCD, S.SchMLRef as MLID "
SQLz = SQLz & "FROM Schedule AS S Inner Join "
SQLz = SQLz & "(Select SGrpSchID, nz(Sum(Strips),0) as Strp, Format(Min(BLCD),'mm/dd/yy') AS mBLCD "
SQLz = SQLz & " FROM(SELECT schgrp.SGrpSchID, DateDiff(""m"",[BLCD],[AssignmentDate]) AS strips, schgrp.BLCD "
SQLz = SQLz & "      FROM SchGrp) as GRP Group BY SGrpSchID) as G ON S.SchID = G.SGrpSchID "
SQLz = SQLz & " WHERE S.SchMLRef =" & SelLease.Column(0) & " ORDER BY CDbl(NumInString(Left(SchNo,3))) DESC, SchNo DESC;"


Me.SelSch.RowSource = SQLz
Me.MLNo = Me.SelLease.Column(1)
Me.SchID = Null
Me.SchNo = Null
Me.BillingSubForm.Requery
Me.GroupList.Requery
End Sub

Private Sub SelSch_Click()
Me.SchID = Me.SelSch.Column(0)
Me.SchNo = Me.SelSch.Column(1)
Me.groupID = Null
Me.GroupList = ""
Me.BillingSubForm.Requery
Me.GroupList.Requery
Me.BillingSubForm.Form.Filter = ""
Me.BillingSubForm.Form.FilterOn = False
End Sub
