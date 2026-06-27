Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST As DAO.Recordset
Dim SQLz As String

Private Sub AddASchedule_Click()
    DoCmd.OpenForm "SchAddf", , , , acFormAdd
    Forms!SchAddf.mlsel = Me.MLID
    'Me.ClientsView_Schedules2.Requery
End Sub

Private Sub Form_Close()

   On Error GoTo Form_Close_Error



   On Error GoTo 0
   Exit Sub

Form_Close_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_ClientsView_Schedules1"
End Sub

Private Sub Form_Load()
    Me.ClientsView_Schedules2.Visible = False
    Me.ScheduleLbl.Visible = False
    Me.MasterLeaseList = Me.MasterLeaseList.ListIndex = (0)
End Sub

Public Sub LPAF_Click()
Dim XLS As String
    Dim Countx As Integer
SQLz = "Select * From Client_LPAF Where ClientId=" & Me.Parent.clientID & ";"
Countx = DCount("clientID", "Client_LPAF", "ClientId=" & Me.Parent.clientID)
If Nz(Countx, 0) <= 0 Then
MsgBox "No Matching Units", vbInformation, "No Result"
Exit Sub
End If
XLS = "|A,BC|B,BC|G,$E|H,$E|J,BC|M,$E|N,BC|O,$E|P,BC|Q,$E|R,$E|s,HH|T,HH|U,HH|"
SendReport2Excel SQLz, "Report", "Report for: " & Me.Parent.ClientName & " All Schedules - Units: " & Countx, True, XLS, Countx
End Sub

Private Sub MasterLeaseList_Click()
    Me.ClientsView_Schedules2.Visible = True
    Me.ScheduleLbl.Visible = True
    Set RST = Me.RecordsetClone
    RST.FindFirst "MLID = " & Me.MasterLeaseList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    SQLz = "Select * FROM SchedulesSubform WHERE MLID =" & Me.MLID & " Order By SCHNo1 DESC, SchID DESC;"
    Me.ClientsView_Schedules2.Form.RecordSource = SQLz
End Sub

Private Sub ScheduleSortForm_Click()
If IsNull(Me.MasterLeaseList) Then Exit Sub
DoCmd.OpenForm "ScheduleSort"
SQLz = "Select * From ScheduleSortOrder Where MLID=" & Me.MLID & ";"
Forms!ScheduleSort.Form.RecordSource = SQLz
End Sub
