Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz, Schx As String



Private Sub RefreshBtn_Click()
    On Error GoTo Err_RefreshBtn_Click

    Dim stDocName As String

    stDocName = "Requery"
    DoCmd.RunMacro stDocName

Exit_RefreshBtn_Click:
    Exit Sub

Err_RefreshBtn_Click:
    MsgBox Err.Description
    Resume Exit_RefreshBtn_Click

End Sub



Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub SetScheduleLabel()
Me.Schedule_Label.caption = "Selected Schedule: " & Me.SchNo
Me.Schedule_Label.BorderColor = RGB(51, 204, 255)
End Sub

Private Sub Detail_Click()
Call SetScheduleLabel
End Sub

Private Sub OpsUnitList_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  '  Dim LeftX As Long
  '  Dim TopY As Long
  '  Me.SchNo.SetFocus
  '  SchX = Me.SchNo
  '  GetCursorPosition LeftX, TopY
  '  ShowControlTip LeftX, TopY, "ControlTipPop", "You're on Schedule: " & SchX
End Sub

Public Sub SchCheck_Click()
    Dim strFileName As String
    Dim ReportX As String
    Dim i As String
    If Forms!ClientsView.[ClientShNm] = "TOYOTA LOGISTICS" Then ReportX = "SchedCheckTLS" Else ReportX = "SchedCheck2"
    DoCmd.OpenReport ReportX, acViewPreview, "", "SchID=" & Me.SchID
    DoCmd.SetWarnings True
End Sub

Public Sub Custodian_Click()
    If Me.SchID & "test" = "test" Then Exit Sub
    Call MakeCustodianFiles(Me.SchID)
End Sub

Public Sub ClientUnitSummary_Click()
Dim XLS As String
    Dim Countx As Integer
SQLz = "Select * From Client_LPAF Where ClientId=" & Me.Client & ";"
Countx = DCount("clientID", "Client_LPAF", "ClientId=" & Me.Client)
If Nz(Countx, 0) <= 0 Then
MsgBox "No Matching Units", vbInformation, "No Result"
Exit Sub
End If
XLS = "|A,BC|B,BC|G,$E|H,$E|J,BC|M,$E|N,BC|O,$E|P,BC|Q,$E|R,$E|s,HH|T,HH|U,HH|"
SendReport2Excel SQLz, "Report", "Report for: " & Me.ClientShNm & " All Schedules - Units: " & Countx, True, XLS, Countx
End Sub
Private Sub Form_Activate()

'    If DLookup("[Access Level]", "TBL_Employees", "[User ID] = FAUserName()") = 2 Then
'        Me.SchedulecheckPDF.Visible = False
'        Me.ScheduleCheckList.Visible = False
'        Me.SerialNum.Visible = False
'        Me.Command76.Visible = False
'        Me.SchNo.Visible = False
'        Me.SchNo1.Visible = False
'
'    Else
'        Me.SchedulecheckPDF.Visible = True
'        Me.ScheduleCheckList.Visible = True
'        Me.SerialNum.Visible = True
'        Me.Command76.Visible = True
'        Me.SchNo.Visible = True
'        Me.SchNo1.Visible = True
'    End If
End Sub


Public Sub SerialNumbers_Click()
Call SetScheduleLabel
    Call MakeWorkTable
    DoCmd.OpenReport "Schedule Serial No", acViewPreview, "", "[SchID]=" & SchID, acNormal
End Sub

Public Sub DocsUnitList_Click()
    Call MakeWorkTable
    Call MakeScheduleReport(Me.SchID)
    'Call MakeDocsReport
   'Calls Function to create DOCS Report.
End Sub
Public Sub OpsUnitList_Click()
Call SetScheduleLabel
    Call MakeWorkTable
    Call MakeOPSReport(, Me.SchID)
    'Calls Function to create OPS Report.
End Sub

Public Sub ScheduleCheckLists_Click()
Call SetScheduleLabel
    Call MakeWorkTable
    DoCmd.OpenReport "ScheduleCheckListNew", acViewPreview, , "SchID=" & Me.SchID, acNormal
End Sub
Private Sub MakeWorkTable()
    DoCmd.SetWarnings False
    SQLz = "Delete * From Schedule_Group_WorkTable;"
    DoCmd.RunSQL (SQLz)
    SQLz = "INSERT INTO Schedule_Group_WorkTable ( ScheduleId ) SELECT " & Me.SchID & " AS SchId;"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub

Private Sub ReportsList_Click()
    Dim FormX, CommandX As String
    FormX = Me.ReportsList.Column(3)
    CommandX = Me.ReportsList.Column(2)
    If FormX = "Local" Then
        CallByName Me, CommandX, VbMethod
    Else
        CallByName FormX, CommandX, VbMethod
    End If
    Me.ReportsList = ""
End Sub

Private Sub SchNo_Click()
    DoCmd.close acForm, "Schedule_Form"
    DoCmd.OpenForm "Schedule_Form", , , "SchID=" & Me.SchID
End Sub
Public Sub View_Active_ATF_Click()
If Me.SchID & "test" = "test" Then Exit Sub
    Dim FAIDx As String
    FAIDx = Nz(DLookup("FAID", "vw_SixKeys", "MLOrig= 'FA' and Schid=" & Me.SchID), "")
    If FAIDx = "" Then
        MsgBox "This NOT an FA Leased Schedule", vbCritical, AtlaasVersion
    Exit Sub
    Else
     Call OpenActiveATFReport(Me.SchID)

    End If
    End Sub
