Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz As String

Private Sub DeleteRecord_Click()
On Error GoTo FixMe
DoCmd.RunSQL ("Delete From UnitTitleLog WHERE UnitTitleLogId=" & Me.UnitTitleLogId & ";")
FixMe:
End Sub
Public Function LoadData(UnitIDx As Long)
SQLz = "SELECT Units.UnitID "
SQLz = SQLz & ", UnitTitleLog.UnitTitleLogDate "
SQLz = SQLz & ", IIf([entrytypeid]=1,'In','Out') AS Status, UnitTitleLog.TitleNumber "
SQLz = SQLz & ", UnitTitleLog.UnitTitleLogId, UnitTitleLog.UnitRef, UnitTitleLog.TitleSt "
SQLz = SQLz & ", UnitTitleLog.TitleSt AS TitleState, UnitTitleLog.EntryReason AS Reason "
SQLz = SQLz & ", IIf([ExpectedReturn]=-1,'Y','N') AS RTN "
SQLz = SQLz & ", UnitTitleLog.UserID "
SQLz = SQLz & "FROM Units INNER JOIN UnitTitleLog ON Units.UnitID = UnitTitleLog.UnitRef "
SQLz = SQLz & "Where Units.UnitID=" & UnitIDx & " AND EntryReasonID <> 41 "
SQLz = SQLz & "ORDER BY UnitTitleLog.UnitTitleLogDate DESC , UnitTitleLog.UnitTitleLogId DESC;"
Me.Form.RecordSource = SQLz
End Function
Private Sub SetLocks()
If FAUserAccess > 24 Then
Me.UnitTitleLogDate.Locked = False
Me.Status.Locked = False
Me.TitleState.Locked = False
Me.DeleteRecord.Visible = True
Else
Me.UnitTitleLogDate.Locked = True
Me.Status.Locked = True
Me.TitleState.Locked = True
Me.DeleteRecord.Visible = False
End If
End Sub

Private Sub Form_Load()
Call SetLocks
End Sub

Private Sub Reason_DblClick(Cancel As Integer)
Call OpenLogEdit
End Sub

Private Sub OpenLogEdit()
If FAUserAccessLevel > 15 Then
DoCmd.close acForm, "OPSTitleLogEntry"
DoCmd.OpenForm "OPSTitleLogEntry", acNormal, , "UnitTitleLogId=" & Me.UnitTitleLogId
End If
End Sub

Private Sub RTN_DblClick(Cancel As Integer)
Call OpenLogEdit
End Sub

Private Sub Status_DblClick(Cancel As Integer)
Call OpenLogEdit
End Sub

Private Sub TitleNumber_DblClick(Cancel As Integer)
Call OpenLogEdit
End Sub

Private Sub UnitTitleLogDate_DblClick(Cancel As Integer)
Call OpenLogEdit
End Sub
