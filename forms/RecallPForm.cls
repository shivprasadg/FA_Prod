Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ExportRecallList_Click()
    Dim msg1, SQLz As String
    SQLz = "SELECT Company, RecallNum, Type, Description, NoticeDate, UnitNumber, UnitVin, LocationName, ActionNeeded, ActionCompleted, EngSerialNo " & _
           "INTO TempRecall From RecallExportReport " & _
           "WHERE Company <> '' "
    If Me.recallno & "test" <> "test" Then SQLz = SQLz & "And RecallNum ='" & Me.recallno.Column(1) & "' "
    If Me.Client & "test" <> "test" Then SQLz = SQLz & "And ClientId =" & Me.Client & " "
    If Me.clientCity & "test" <> "test" Then SQLz = SQLz & "And LocationID =" & Me.clientCity & " "
    SQLz = SQLz & " Order By NoticeDate Desc, LocationName ;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    FormatExcel "TempRecall", "C:\Faas\" & FAUserName & "\Temp\RecallExportReport", "Recall", "Close", , DCount("Company", "Recall Export"), Me.Client & " Recall " & Date & " - " & DCount("Company", "TempRecall") & " Units Listed", True
    DoCmd.DeleteObject acTable, "TempRecall"
    DoCmd.SetWarnings True
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'><b>Recall List</b></p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:15px'>"
    msg1 = msg1 & "Subject:<b>&#160;&#160;&#160;&#160;&#160;Recall List RE: " & Me.Client.Column(1) & " - Recall: " & Me.recallno.Column(1) & "</b><br />"
    msg1 = msg1 & "Priority:<b>&#160;&#160;&#160;&#160;&#160;&#160;High </b><br/><br/>"
    msg1 = msg1 & "Dear  <br/><br/>The following units are part of a recall notice<br /><br />"
    msg1 = msg1 & "Client: <B>" & Me.Client.Column(1) & "</B><br />"
    msg1 = msg1 & "Please see the attached vehicle list:"
    msg1 = msg1 & " " & Msgx & "</P>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:15px'><br/>"
    msg1 = msg1 & "Regards,<br /><br /><br /><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:9px'><br/>[End Msg]"
    msg1 = msg1 & "Email Created: " & Date & " (via Fleet Advantage Data Manager).<br/><br />"
    msg1 = msg1 & "</body></html>"
    Call SendEmailMessage("", "Re: " & Me.Client.Column(1) & " - Recall: " & Me.recallno.Column(1), msg1, False, True, False, "", , , "C:\Faas\" & FAUserName & "\Temp\RecallExportReport.xlsx")
End Sub

Private Sub client_Click()
Me.clientCity = Null
Me.clientCity.Requery
End Sub
Private Sub recallno_Click()
Me.Client = Null
Me.Client.Requery
Me.clientCity = Null
Me.clientCity.Requery
End Sub
