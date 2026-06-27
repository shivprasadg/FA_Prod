Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz, SQLx, GroupIDx As String
Dim Ccrl       As Control
Private Sub Prev_Click()
 '   Call fieldname_Click
End Sub
Private Sub New_Click()
 '   Call fieldname_Click
End Sub
Private Sub ActionDate_Click()
  '  Call fieldname_Click
End Sub
Private Sub ActionTaken_Click()
  '  Call fieldname_Click
End Sub

Private Sub Send_EmailAttachement_Click()
'Call SendATFEmail(Me.SchID, Me.groupID, Me.SelectedID)
End Sub


Private Sub User_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_SysUser = Me.User
    Call SetRecordSource(" SysUser = '" & Me.User & "' ")
End Sub

Private Sub ActionDate_DblClick(Cancel As Integer)
  '  Call ClearSearchBoxes("All")
  '  Me.Search_ActionDATE = SplitString(Me.ActionDate, " ", 0)
  '  Call SetRecordSource("ActionDate Like '*" & Me.Search_ActionDATE & "*' ")
End Sub

Private Sub ClearResetForm_Click()
    Call ClearLabel_Click
    ' Me.Form.RecordSource = "SELECT * From vw_AuditUnits Order By ActionDate Desc"
End Sub

Private Sub ClearLabel_Click()
  '  On Error Resume Next
  '  Me.Parent.UnitRevDetail.RowSource = ""
  '  Me.Parent.UnitRevDetail.Visible = False
    On Error GoTo 0
    Call ClearSearchBoxes("All")
    Call SetRecordSource
End Sub

Private Sub FieldName_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    'Me.Search_FieldName = Me.FieldName
    Call SetRecordSource(" ColumnName = '" & Me.FieldName & "' ")
End Sub
Public Function LoadData(UnitIDx As Long)
    On Error Resume Next
    If Nz(Me.AssetID) = 0 Or Me.AssetID <> UnitIDx Then
        SQLz = "SELECT * FROM vw_AuditUnits "
        SQLz = SQLz & "Where AssetID =" & UnitIDx & " "
        SQLz = SQLz & "ORDER BY vw_AuditUnits.ActionDate DESC;"
    End If
    On Error GoTo 0
    Me.Form.RecordSource = SQLz
    Call GetCount
End Function
Public Sub Form_Current()
On Error Resume Next
    If Me.AssetID <> Me.Parent.UnitID Then Call LoadData(Me.AssetID)
On Error GoTo 0
  ' GroupIDx = Me.groupID
 ' Me.SelectedID = Me.AssetID
 '  Me.txtCurrent = Me.FieldName
End Sub
Public Function SetRecordSource(Optional WhereX As String)
  '  On Error GoTo Fixme
   ' SQLz = ""
   ' SQLz = SQLz & "SELECT * From vw_AuditUnits "
   ' If WhereX <> "" Then SQLz = SQLz & "WHERE " & WhereX & " "
   ' SQLz = SQLz & "Order By ActionDate Desc;"
   ' Me.Form.RecordSource = SQLz

FixMe:
End Function
Private Sub Form_Load()
    Call ClearSearchBoxes("All")
  '  Call GetCount
    On Error GoTo 0
End Sub


Private Sub Search_FieldName_Click()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
'    Call SetRecordSource("AND FieldName = '" & Me.Search_FieldName & "' ")
End Sub

Private Sub Search_New_AfterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
'    Call SetRecordSource("AND NewValue Like '*" & Me.Search_New & "*' ")
End Sub
Private Sub Search_Prev_AfterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
'    Call SetRecordSource("AND OldValue Like '*" & Me.Search_Prev & "*' ")
End Sub
Private Sub Search_SysUser_Click()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    If Me.Search_SysUser.Text = "" Then
        Me.Form.RecordSource = "SELECT * From vw_AuditUnits"
    Else
        Me.Form.RecordSource = "SELECT * From vw_AuditUnits Where SysUser='" & Me.Search_SysUser & "'"
    End If
End Sub

Private Sub Search_ActionDate_afterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
'    Call SetRecordSource("AND ActionDate Like '*" & Me.Search_ActionDATE & "*' ")
End Sub

Private Sub ClearSearchBoxes(Excludeme As String)
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Search" And Me(Ccrl.Name).Name <> Excludeme Then
            Me(Ccrl.Name) = ""
        End If
    Next Ccrl
End Sub

Private Sub GetCount()
    Dim RST   As DAO.Recordset
    Set RST = Me.RecordsetClone
    On Error Resume Next
    RST.MoveLast
    Me.Events.Value = RST.RecordCount
    RST.close
    Set RST = Nothing
    On Error GoTo 0
End Sub
