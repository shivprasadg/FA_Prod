Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz, SQLx, GroupIDx As String
Dim sFieldName As String
Dim Ccrl       As Control
Private Sub Prev_Click()
    Call fieldname_Click
End Sub
Private Sub New_Click()
    Call fieldname_Click
End Sub
Private Sub ActionDate_Click()
    Call fieldname_Click
End Sub
Private Sub ActionTaken_Click()
    Call fieldname_Click
End Sub
Private Sub User_Click()
    Call fieldname_Click
End Sub

Private Sub ActionDate_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_ActionDate = SplitString(Me.ActionDate, " ", 0)
    Call SetRecordSource("AND ActionDate Like '*" & Me.Search_ActionDate & "*' ")
End Sub

Private Sub ClearResetForm_Click()
    Call ClearLabel_Click
End Sub

Private Sub ClearLabel_Click()
    Call ClearSearchBoxes("All")
    Call SetRecordSource
End Sub
Public Sub HideDetails()
    Me.Parent.UnitRevDetail.RowSource = ""
    Me.Parent.UnitRevDetail.Visible = False
    Me.Parent.ReasonAction.Visible = False
    Me.Parent.ReasonNote.Visible = False
  '  Me.Parent.UnitRevDetailLabel.caption = Me.FieldName & " Detail"
   ' Me.Parent.ReasonAction = ""
   ' Me.Parent.ReasonNote = ""
   ' Me.Parent.ReasonINDX = ""
End Sub
Private Sub FieldName_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_FieldName = Me.FieldName
    Call SetRecordSource("AND FieldName = '" & Me.Search_FieldName & "' ")
End Sub
Private Sub fieldname_Click()
    If Me.Parent.UnitID & "test" <> "test" Then Exit Sub

    On Error GoTo 0
    Call HideDetails

    If Me.FieldGroup = "UnitSpec" _
       Or Me.FieldGroup = "Rent" _
       Or Me.FieldGroup = "Residual" _
       Or Me.FieldGroup = "Vendor" _
       Or Me.FieldGroup = "Economics" _
       Or Me.FieldGroup = "Structure" _
       Or Me.FieldGroup = "Terms" Then

        Dim FieldGroupX, DateTimeStart, DateTimeEnd, JustDate, ViewX As String
        Dim UnitIDStartLng As Long
        JustDate = SplitString(Me.ActionDate, " ", 0)
        If Len(JustDate) < 10 Then JustDate = "0" & JustDate
        If Len(JustDate) < 10 Then JustDate = Left(JustDate, 3) & "0" & Right(JustDate, 6)
        SQLx = "SELECT AssetID, Ccur(OldValue) As Prev, Ccur(NewValue) As New, '" & Me.FieldName & "' As FieldName, sysUser, ActionDate, GroupID "

        If Me.FieldName Like "*Rent*" Or Me.FieldName Like "*Cost*" Or Me.FieldName Like "*Resi*" Then
            If Me.FieldName Like "*Rent*" Or Me.FieldName Like "*Cost*" Then
                FieldGroupX = "Rent Change"
                ViewX = "vw_FAID_RevRents"
                If Me.Parent.GroupID & "t" <> "t" Then UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.GroupID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            End If
        End If

        If Me.FieldName Like "*Resi*" Then
            FieldGroupX = "Residual Change"
            ViewX = "vw_FAID_RevResidual"
            If Me.Parent.GroupID & "t" <> "t" Then UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.GroupID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
        End If

        If Me.FieldName Like "*Vendor*" Or Me.FieldName Like "*delivery*" Then
            FieldGroupX = "Vendor Change"
            ViewX = "vw_FAID_RevUnitVendor"
            If Me.Parent.GroupID & "t" <> "t" Then UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.GroupID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            SQLx = "SELECT AssetID, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldName, sysUser, ActionDate, GroupID "
        End If

        If Me.FieldGroup = "UnitSpec" Then
            FieldGroupX = "Unit Spec Change"
            ViewX = "vw_FAID_RevUnitSpec"
            If Me.Parent.GroupID & "t" <> "t" Then UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.GroupID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            SQLx = "SELECT AssetID, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldName, sysUser, ActionDate, GroupID "
        End If
        If Me.FieldGroup = "Structure" Then
            FieldGroupX = "Structure"
            ViewX = "vw_FAID_RevMoveGrps"
            If Me.Parent.GroupID & "t" <> "t" Then UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.GroupID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            SQLx = "SELECT AssetID, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldName, sysUser, ActionDate, GroupID "
        End If
        DateTimeStart = Left(Me.ActionDate, Len(Me.ActionDate) - 5) & "00 " & Right(Me.ActionDate, 2)
        DateTimeEnd = Left(Me.ActionDate, Len(Me.ActionDate) - 5) & "02 " & Right(Me.ActionDate, 2)
        Me.Parent.UnitRevDetail.Visible = True
        Me.Parent.ReasonAction.Visible = True
        Me.Parent.ReasonNote.Visible = True
        Me.Parent.UnitRevDetail.ColumnWidths = "0.6in;1.5in;1.5in;0in;0;1in;0"
        'RealFieldX = DLookup("FAIDField", "ATF_FAID_FieldMatrix", "FieldLabel='" & Me.FieldName & "'")

        SQLx = SQLx & "FROM  " & ViewX & " " & _
               "WHERE GroupID=" & Me.Parent.GroupID & " AND FieldName='" & Me.FieldName & "' " & _
               "AND DateValue(ActionDate) =  DateValue('" & Me.ActionDate & "') Order By AssetID, ActionDate Desc;"

        If Me.FieldGroup = "Terms" Then
            FieldGroupX = "Terms Change"
            ViewX = "vw_FAID_RevTerms"
            ' UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.SGrpID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            SQLx = "SELECT FieldName, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldType, sysUser, ActionDate, GroupID " & _
                   "FROM  vw_FAID_RevTerms " & _
                   "WHERE GroupID=" & Me.Parent.GroupID & " " & _
                   "AND DateValue(ActionDate) =  DateValue('" & Me.ActionDate & "') Order By ActionDate Desc;"
            Me.Parent.UnitRevDetail.ColumnWidths = "1.3in;1.0in;1.0in;0in;0;1in;0"
        End If
        If Me.FieldName Like "*Swaps*" Then
            Dim EcoType As String
            ViewX = "vw_FAID_RevEconomics"
            If Me.FieldName = "Sales Swaps" Then EcoType = "FieldName Like '*(sales)*'"
            If Me.FieldName = "Synd Swaps" Then EcoType = "Not FieldName Like '*(sales)*'"
            SQLx = "SELECT FieldName, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldType, sysUser, ActionDate, GroupID " & _
                   "FROM  vw_FAID_RevEconomics " & _
                   "WHERE GroupID=" & Me.Parent.GroupID & " AND " & EcoType & " " & _
                   "AND DateValue(ActionDate) =  DateValue('" & Me.ActionDate & "') Order By ActionDate Desc;"
            Me.Parent.UnitRevDetail.ColumnWidths = "1.3in;1.0in;1.0in;0in;0;1in;0"
        End If

        Me.Parent.UnitRevDetail.RowSource = SQLx
        SQLx = ""
GetChangeReason:
        SQLx = SQLx & "Select * From ATF_ChangeReasons "
        SQLx = SQLx & "WHERE CDate(SplitString(ChangeDate,' ',0))='" & CDate(JustDate) & "' "
        SQLx = SQLx & "AND [FieldName]='" & Me.FieldName & "' "
        'Check if this was a 'unit level field' event
        If Me.FieldGroup <> "Economics" And Me.FieldGroup <> "Terms" Then
            SQLx = SQLx & "AND [UnitID]='" & UnitIDStartLng & "'"
        Else
            SQLx = SQLx & ";"
        End If
        Dim RST As DAO.Recordset
        On Error GoTo FixMe
        Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
        On Error Resume Next
        If RST.RecordCount < 1 Then Exit Sub
        RST.MoveLast
        RST.MoveFirst
        If RST.RecordCount < 1 Then Exit Sub
        On Error GoTo 0
        RST.MoveFirst
        Me.Parent.ReasonAction = RST!Action
        Me.Parent.ReasonNote = RST!Reason
        Me.Parent.ReasonINDX = RST!INDX
        RST.close
        Set RST = Nothing
    End If
Exit Sub
FixMe:
Set RST = Nothing
'MsgBox "There was a SQL on this event", vbCritical, AtlaasVersion
End Sub
Public Sub Form_Current()
 ' GroupIDx = Me.groupID
 ' Me.SelectedID = Me.identity
'  Me.txtCurrent = Me.FieldName
End Sub
Public Function SetRecordSource(Optional WhereX As String)
    On Error GoTo FixMe
    Dim xSql As String
    xSql = Me.Form.RecordSource
    xSql = SplitString(xSql, " Order By", 0)


    If Me.Parent.UnitID & "t" <> "t" And WhereX Like "*FieldName*" Then WhereX = Replace(WhereX, "FieldName", "FieldLabel")
    If Me.Parent.UnitID & "t" <> "t" And WhereX Like "*Action*" Then WhereX = Replace(WhereX, "ActionDesc", "[Action]")
    If WhereX = "" And xSql Like "*AND*" Then xSql = SplitString(xSql, "AND", 0)
    SQLz = ""
    SQLz = xSql & " " & WhereX & " ORDER BY ActionDate DESC;"
    Me.Form.RecordSource = SQLz
    Call GetCount
FixMe:
End Function
Private Sub Form_Load()
    On Error Resume Next
    On Error GoTo 0
    Call ClearSearchBoxes("All")
End Sub

Private Sub Search_ActionTaken_AfterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND ActionDesc Like '*" & Me.Search_ActionTaken & "*' ")
End Sub

Private Sub Search_FieldName_Click()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND FieldName = '" & Me.Search_FieldName & "' ")
End Sub

Private Sub Search_New_AfterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND NewValue Like '*" & Me.Search_New & "*' ")
End Sub
Private Sub Search_Prev_AfterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND OldValue Like '*" & Me.Search_Prev & "*' ")
End Sub
Private Sub Search_SysUser_Click()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND SysUser =    '" & Me.Search_SysUser & "' ")
End Sub

Private Sub Search_ActionDate_afterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND ActionDate Like '*" & Me.Search_ActionDate & "*' ")
End Sub

Private Sub ClearSearchBoxes(Excludeme As String)
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Search" And Me(Ccrl.Name).Name <> Excludeme Then
            Me(Ccrl.Name) = ""
        End If
    Next Ccrl
End Sub


Private Sub User_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_SysUser = Me.User
    Call SetRecordSource("AND SysUser =    '" & Me.Search_SysUser & "' ")
End Sub

Public Sub GetCount()
    Dim RST    As Object
    If Me.Parent.SchID <> "" Or Me.Parent.GroupID <> "" Then
        Set RST = Me.RecordsetClone
        On Error Resume Next
        RST.MoveLast
        On Error GoTo 0
        Me.Events.Value = RST.RecordCount
        RST.close
        Set RST = Nothing
    End If
End Sub
