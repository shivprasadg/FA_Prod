Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Askit, dTypex, FilterX, SQLy As String
Dim FilterChange As Boolean



Private Sub DeleteRecord_Click()
    On Error GoTo FixMe
    Askit = MsgBox("You are about to Delete an Document Link!" & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Kill it?")
    If Askit <> 6 Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From DocLinks WHERE DocId=" & Me.DocID & ";")
    DoCmd.SetWarnings True
    Me.Form.Requery
FixMe:
End Sub


Private Sub EditLink_Click()
    DoCmd.close acForm, "DocAddEditLinkForm"
    DoCmd.OpenForm "DocAddEditLinkForm", , , "DocID=" & Me.DocID, , , "|||||" & Me.Parent.Form.Name & "|Edit"
End Sub

Private Sub Form_Current()
    On Error Resume Next
    Me.txtCurrent = Me.[DocID]
    On Error GoTo 0
    'If FilterChange = True Then Call GetCount
End Sub

Public Sub GetCount()
    Dim rs As DAO.Recordset
    'Me.Form.Requery
    Set rs = Me.Recordset.Clone
    rs.MoveLast
    rs.MoveFirst
    Me.UnitCount = rs.RecordCount
    FilterChange = False
End Sub
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    If Not IsNull(Me.Filter) Then FilterX = Me.Filter Else FilterX = ""
    If FilterX <> Me.Filter Then FilterChange = True Else FilterChange = False
End Sub

Private Sub Form_Load()
'Call GetCount
    FilterChange = False
    SQLy = "SELECT * FROM DocLinks "
    If Me.Parent.Form.Name = "GroupForm" Then SQLy = SQLy & " Where GroupId = " & Me.Parent.GroupID
    If Me.Parent.Form.Name = "Schedule_Form" Then SQLy = SQLy & " Where SchId = " & Me.Parent.SchID
    If Me.Parent.Form.Name = "MstrLease" Then SQLy = SQLy & " Where MLID = " & Me.Parent.MLID
    If Me.Parent.Form.Name = "LinkedDocsManager" Then SQLy = SQLy & " Where SCHID = " & Me.Parent.SchID
    Me.Form.RecordSource = SQLy
End Sub
