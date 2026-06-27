Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim xSql       As String
Dim RST        As DAO.Recordset
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub ClearClientSearch_Click()
    Me.SearchClient = ""
    Call SearchClient_Change
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    Me.SchNo = "Please select a client..."
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
        "From vw_SixKeys " & _
        "GROUP BY clientID, ClientShortName, CompanyName " & _
        "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz
    Me.LeaseList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.DocLinksSub.Form.Visible = False
End Sub


Public Sub ClientList_Click()
    SQLz = "SELECT MLID, MLNo, MLOrig, ClientID From vw_SixKeys Where MLOrig='FA' " & _
        "GROUP BY MLID, MLNo, MLOrig, ClientID Having ClientID =" & Me.ClientList & " ORDER BY MLOrig;"
    Me.LeaseList.RowSource = SQLz
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.SchID = Null
    Me.groupID = Null
    Me.SchNo = Me.ClientList.Column(1)
    Me.DocLinksSub.Form.RecordSource = ""
    Me.DocLinksSub.Form.Visible = False
End Sub
Public Sub Leaselist_Click()
    SQLz = "SELECT S.SchID, Schedule, S.MLID, Count(AssetID) AS Unt, DType, S.ClientID, MLNo, MLOrig " & _
        "FROM vw_SixKeys AS S " & _
        "Left Join (SELECT SchID, DType FROM DocLinks WHERE DType = 'DIR') As D ON S.SchID = D.SchID  " & _
        "WHERE MLOrig='FA' " & _
        "GROUP BY S.SchID, Schedule, S.MLID, S.ClientID, MLNo, DType, MLOrig HAVING MLID=" & Me.LeaseList & " " & _
        "ORDER BY CDbl(NumInString(Left(Schedule,3))) Desc,  Schedule Desc;"
    Me.ScheduleList.RowSource = SQLz
    Me.SchID = Null

    Me.groupID = Null
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1)
    Me.GroupList.RowSource = ""
    Me.DocLinksSub.Form.RecordSource = ""
    Me.DocLinksSub.Form.Visible = False
   End Sub

Public Sub ScheduleList_Click()
    Me.SchID = Me.ScheduleList
    Me.groupID = Null
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(1)
    SQLz = "SELECT SgrpID, UnitGroup, Count(AssetID) AS Unt, SchID " & _
        "FROM vw_SixKeys Where MLOrig='FA' " & _
        "GROUP BY SgrpID, UnitGroup,SchID HAVING SchID=" & Me.ScheduleList & " ORDER BY UnitGroup;"
    Me.GroupList.RowSource = SQLz
    SQLz = ""
    SQLz = "SELECT * FROM DocLinks WHERE SchId=" & Me.SchID & ";"
    Me.DocLinksSub.Form.RecordSource = SQLz
    Me.DocLinksSub.Form.Visible = True
End Sub
Public Sub GroupList_Click()
    Me.groupID = Me.GroupList.Column(0)
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(1) & " Grp: " & GroupList.Column(1)
        SQLz = "SELECT * FROM DocLinks WHERE SchId=" & Me.groupID & ";"
    Me.DocLinksSub.Form.RecordSource = SQLz
End Sub


Private Sub SearchClient_Change()
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
        "From vw_SixKeys WHERE ClientShortName Like '*" & Me.SearchClient.Text & "*' Or CompanyName Like '*" & Me.SearchClient.Text & "*' " & _
        "GROUP BY clientID, ClientShortName, CompanyName " & _
        "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz
End Sub


Private Sub AddDocLink_Click()
    DoCmd.close acForm, "DocAddEditLinkForm"
    DoCmd.OpenForm "DocAddEditLinkForm", acNormal, , , acFormAdd, , Me.ClientList & "|" & Me.LeaseList & "|" & Me.SchID & "|0|0|" & Me.Form.Name & "|Add"
End Sub
