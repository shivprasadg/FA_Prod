Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset
Dim SQLz       As String

Private Sub AddNew_Click()
    DoCmd.GotoRecord , , acNewRec
End Sub

Private Sub CoLesseeID_Click()
    Me.CoLesseeTaxID = Me.CoLesseeID.Column(2)
    If Me.Dirty Then
        Me.Dirty = False
        Me.CoLesseeList.Requery
    End If
End Sub

Private Sub CoLesseeLabel_Click()
    If Me.Dirty Then
        'Me.Dirty = False
        Me.CoLesseeList.Requery
    End If
End Sub

Private Sub CoLesseeList_Click()
    Set RST = Me.RecordsetClone
    RST.FindFirst "INDX = " & Me.CoLesseeList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
End Sub

Private Sub DeleteMe_Click()
    Dim ResposeX As VbMsgBoxResultEx
    Dim msg1   As String
    msg1 = msg1 & "<b><Font Size=3>Delete!</font></b><br/>"
    msg1 = msg1 & "You are about to delete the <b> " & Me.CoLesseeLabel & "<br/>" & CoLesseeID.Column(1) & "</b><br /><br/>"
    msg1 = msg1 & "  <b>Are you sure?</b><br/>"
    msg1 = msg1 & "  Please select..."
    ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
    If ResposeX = vbBt2 Then Exit Sub
    If ResposeX = vbBt1 Then
        SQLz = "Delete FROM ScheduleCoLessee Where INDX = " & Me.INDX
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.CoLesseeList.Requery
    End If
End Sub

Private Sub Form_Current()
    SQLz = ""
    SQLz = SQLz & "SELECT ClientID, clientCompanyName AS CoLessee, clientTaxID As TaxID, ClientGroupId "
    SQLz = SQLz & "FROM  Clients "
    SQLz = SQLz & "WHERE ClientGroupID= " & Me.Parent.ClientGroupID & " "
    SQLz = SQLz & "ORDER BY clientCompanyName"
    Me.CoLesseeID.RowSource = SQLz
    SQLz = ""
    SQLz = SQLz & "SELECT Sch.INDX, Sch.CoLesseeLabel, Clients.clientCompanyName, Sch.ScheduleId, Clients.clientID "
    SQLz = SQLz & "FROM ScheduleCoLessee As Sch "
    SQLz = SQLz & "INNER JOIN Clients ON Sch.CoLesseeID = Clients.clientID WHERE Sch.ScheduleId=" & Me.Parent.SchID & " "
    SQLz = SQLz & "ORDER BY Clients.clientCompanyName;"
    Me.CoLesseeList.RowSource = SQLz
End Sub


Private Sub SaveCoLessee_Click()
    If Me.Dirty Then
        Me.Dirty = False
        Me.CoLesseeList.Requery
    End If

End Sub
