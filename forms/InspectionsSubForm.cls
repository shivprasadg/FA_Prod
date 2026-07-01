Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz As String
Dim Committ As String

Private Sub AddInspection_Click()
    Dim Notex, PrevLink As String
    Dim xNum As Integer
    xNum = Nz(DMax("InspNum", "UnitInspections", "AssetID= " & Me.Parent.UnitID), 0) + 1
    Committ = MsgBox("You are about to add in Inspection Num " & xNum & vbNewLine & "Do you agree?", vbYesNo + vbQuestion, "Adding a Inspection")
    If Committ = vbNo Then Exit Sub
    PrevLink = Nz(DLookup("iWeblink", "UnitInspections", "Not iWeblink is Null and AssetID=" & Forms!Unitsform.Form.UnitID), "")
    Notex = "This Inspection Num (" & xNum & ") was manually added in by: " & FaUserName & " on: " & Date & ""
    If PrevLink <> "" Then Notex = Notex & vbNewLine & "Note: - the web link is from a previous inspection"
    SQLz = "Insert Into UnitInspections (AssetID, InspNum,RequestedBy, iResultID, iVendorID, iDate, iTypeID, InvoiceNumber, iNotes2, iWeblink) "
    SQLz = SQLz & "Select " & Forms!Unitsform.Form.UnitID & "," & xNum & ",'" & FaUserName & "', 3, 253,'" & Date & "', 4 "
    SQLz = SQLz & ", 'FA-" & Me.InspnID & "'"
    SQLz = SQLz & ",'" & Notex & "', '" & PrevLink & "'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call LoadData(Forms!Unitsform.Form.UnitID)
    Me.Form.Refresh
End Sub

Private Sub DeleteInsp_Click()
    Dim iNum As Integer
    Committ = MsgBox("You are about to Delete Inspection Num " & Me.InspNum & vbNewLine & "Are you SURE?", vbYesNo + vbCritical, "DELETE!!!")
    If Committ = vbNo Then Exit Sub
    DoCmd.RunSQL ("Delete From UnitInspections Where InspnID=" & Me.InspnID & ";")
    Me.InspectionList.Requery
    Me.Form.Requery
End Sub

Public Function LoadData(UnitIDx As Long)
    Me.Form.RecordSource = "Select * From UnitInspections WHERE AssetID=" & UnitIDx & " ORDER BY iDate DESC"
    If Me.iResultID = 1 Then Me.iPassedDate.Enabled = True Else Me.iPassedDate.Enabled = False
    SQLz = "SELECT I.InspnID, I.InspNum As Insp"
    SQLz = SQLz & ", I.iDate AS [Date]"
    SQLz = SQLz & ", R.InspectionResult AS Result"
    SQLz = SQLz & ", T.InspectionType AS Type "
    SQLz = SQLz & " FROM (UnitInspections as I INNER JOIN UnitInspectionResults AS R ON I.iResultID = R.InpsectionResultId) "
    SQLz = SQLz & " INNER JOIN UnitInspectionTypes as T ON I.iTypeID = T.InspectionTypeId "
    SQLz = SQLz & " WHERE AssetID=" & UnitIDx & " ORDER BY I.iDate DESC;"
    Me.InspectionList.RowSource = SQLz
    Me.InspnID.SetFocus
    DoCmd.GotoRecord , , acFirst
End Function

Private Sub Form_Current()
    On Error Resume Next
    If Me.AssetID <> Me.Parent.UnitID Then Call LoadData(Me.AssetID)
    On Error GoTo 0
End Sub

Private Sub InspectionList_Click()
    Me.Form.FilterOn = False

    Set RST = Me.RecordsetClone
    RST.FindFirst "InspnID = " & Me.InspectionList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    Set RST = Nothing
End Sub

Private Sub iPassedDate_AfterUpdate()
    If IsDate(iPassedDate) Then Call SetNetBuyOut(True)
    Call AutoUpdateEOL_Statuses(True)
End Sub

Private Sub iResultID_Change()
    If Me.iResultID = 1 Then
        Me.iPassedDate.Enabled = True
    Else
        Me.iPassedDate = Null
        Me.iPassedDate.Enabled = False
    End If
    Call AutoUpdateEOL_Statuses(True)
End Sub
