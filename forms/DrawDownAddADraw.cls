Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST As DAO.Recordset
Dim SQLz   As String

Private Sub AddDrawBtn_Click()
Call LoadData("Add", Me.groupID)
End Sub

Private Sub CloseMe_Click()
  ' Call SaveMe_Click
   If IsOpen("DrawDownManager") Then Forms!DrawDownManager.Form.DrawDownList.Requery
   DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub DrawBankId_AfterUpdate()
    'Dim Committ As Variable
If IsNull(Me.DrawDownId) Then Exit Sub
   ' Committ = MsgBox("Do you want to set " & Me.DrawBankId.Column(1) & " as the Interim Bank" & _
   '     "" & vbNewLine & "for All units/parts in " & _
   '     "" & vbNewLine & "Draw: " & Me.DrawName & " ?", vbInformation + vbYesNo, "Set InterimBank")
   ' If Committ = vbNo Then Exit Sub
   ' If Committ = vbYes Then
        SQLz = "Update Units Set UnitIFAmt=Nz(VendorCost,UnitFinalCost), UnitIFBankID=" & Me.DrawBankId.Column(0) & " Where DrawDownId=" & Me.DrawDownId & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        SQLz = "Update ChildParts Set cIFAmount=childFinalCost, cIFBankID=" & Me.DrawBankId.Column(0) & " Where DrawDownId=" & Me.DrawDownId & ";"
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
   ' End If
End Sub

Private Sub DrawBankId_Click()

End Sub

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_UnitNotes"
End Sub
Public Function LoadData(TaskType As String, IDtoUSEx As Long)

    SQLz = "SELECT Banks.BankID, Banks.BankShortName "
    SQLz = SQLz & "FROM Banks "
    SQLz = SQLz & "ORDER BY Banks.BankShortName;"
    Me.DrawBankId.RowSource = SQLz
    If TaskType = "ADD" Then
        Me.Header.caption = "Draw Down Add"
        SQLz = "SELECT clientID,ClientShortName,MlNo,MLID,FAID,Schedule,UnitGroup,FundedDate,BankShortName,GroupID, SchID "
        SQLz = SQLz & "FROM vw_SixKeys "
        SQLz = SQLz & "WHERE sGrpId=" & IDtoUSEx & " "
        SQLz = SQLz & "GROUP BY clientID,ClientShortName,MlNo,MLID,FAID,Schedule,UnitGroup,FundedDate,BankShortName,GroupID,SchID;"
        Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
        On Error Resume Next
        RST.MoveFirst
        On Error GoTo 0
        DoCmd.GotoRecord , , acNewRec
        Me.DrawNumber = CInt(Nz(DMax("DrawNumber", "DrawDowns", "GroupID=" & IDtoUSEx), 0)) + 1
        Me.DrawName = "Sch:" & RST!Schedule & "|Grp:" & RST!UnitGroup & "|Draw:" & Me.DrawNumber & ""
        Me.DrawTypeId = 1
        Me.DrawBankId = 5649
        Me.Client = RST!ClientShortName
        Me.Assignee = RST!BankShortName
        Me.FundDate = Nz(RST!FundedDate, "none")
        Me.CreatedBy = FAUserName
        Me.CreatedOn = Date
        Me.clientID = RST!clientID
        Me.MLID = RST!MLID
        Me.SchID = RST!SchID
        Me.groupID = RST!groupID
        Me.FAID = RST!FAID
        RST.close
        Set RST = Nothing
    End If
    If TaskType = "Edit" Then
        Me.Header.caption = "Draw Down Edit"
        SQLz = "Select * From DrawDowns Where DrawDownID=" & IDtoUSEx
        Me.Form.RecordSource = SQLz
        Me.DrawNumber.ControlSource = "DrawNumber"
        Me.DrawName.ControlSource = "DrawName"
        Me.DrawTypeId.ControlSource = "DrawTypeId"
        Me.CreatedBy.ControlSource = "CreatedBy"
        Me.CreatedOn.ControlSource = "CreatedOn"
        Me.DrawBankId.ControlSource = "DrawBankId"
        Me.DateDrawn.ControlSource = "DateDrawn"
        Me.DrawDownId.ControlSource = "DrawDownId"
        Me.clientID.ControlSource = "clientID"
        Me.MLID.ControlSource = "MLID"
        Me.SchID.ControlSource = "SchID"
        Me.groupID.ControlSource = "groupID"
        Me.FAID.ControlSource = "FAID"
        Me.DrawRequestNumber.ControlSource = "DrawRequestNumber"
        Me.DrawBankAccountNumber.ControlSource = "DrawBankAccountNumber"
        Me.DrawCreditLimit.ControlSource = "DrawCreditLimit"
        Me.DrawPaidOff.ControlSource = "DrawPaidOff"
        Me.Client = DLookup("ClientShNm", "Clients", "ClientID=" & Me.clientID)
        Me.FundDate = DLookup("FundDate", "SchGrp", "SGrpId=" & Me.groupID)
        Me.Assignee = DLookup("Assignee_Bank", "SchGrp", "SGrpId=" & Me.groupID)
    End If

End Function

Private Sub Form_Load()
  If Nz(Me.OpenArgs, "") <> "" Then Call LoadData(SplitString(Me.OpenArgs, ">", 0), SplitString(Me.OpenArgs, ">", 1))
End Sub

Private Sub SaveMe_Click()
    Dim msg As String
    If Me.DrawBankId & "test" = "test" Then msg = msg & "Missing Interim Bank"
    If Me.DrawName & "test" = "test" Then msg = msg & "Missing Draw Name Bank"
    If msg <> "" Then

        Call MsgBox("You cannot SAVE this record:" & vbNewLine & vbNewLine & msg & _
            "" & vbNewLine & vbNewLine & "Nothing Saved" & _
            "" & vbNewLine & vbNewLine & "Do you want to continue without Saving?", vbCritical + vbYesNo, "Missing Data")
        If Committ = vbNo Then Exit Sub
        If Committ = vbYes Then
            Me.Undo
            GoTo continue
        End If
    Else
        If Me.Header.caption = "Draw Down Add" Then
            SQLz = "INSERT INTO DrawDowns ("
            SQLz = SQLz & "MLID,clientID,DrawBankId,DrawName"
            SQLz = SQLz & ",DrawNumber,DrawBankAccountNumber"
            SQLz = SQLz & ",DrawRequestNumber,FAID"
            SQLz = SQLz & ",DrawCreditLimit,CreatedOn"
            SQLz = SQLz & ",CreatedBy,DrawTypeId"
            SQLz = SQLz & ",groupID,SchID) "
            SQLz = SQLz & "Select "
            SQLz = SQLz & Me.MLID & "," & Me.clientID & "," & Me.DrawBankId & ",'" & Me.DrawName & "' "
            SQLz = SQLz & "," & Me.DrawNumber & ",'" & Me.DrawBankAccountNumber & "'"
            SQLz = SQLz & ",'" & Me.DrawRequestNumber & "','" & Me.FAID & "'"
            SQLz = SQLz & "," & Nz(Me.DrawCreditLimit, 0) & ",'" & Me.CreatedOn & "'"
            SQLz = SQLz & ",'" & Me.CreatedBy & "'," & Me.DrawTypeId
            SQLz = SQLz & "," & Me.groupID & "," & Me.SchID
            SQLz = SQLz & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            Call MsgBox("Draw: " & Me.DrawName & " has been saved" & venewline & vbNewLine & "The form will Clear", vbInformation, "Draw Saved")
            Dim Ccrl   As Control
            For Each Ccrl In Me.Controls
                If Me(Ccrl.Name).Tag = "Ditem" Then Me(Ccrl.Name) = Null
            Next Ccrl
        End If
        If Me.Header.caption = "Draw Down Edit" Then
           Me.Form.Refresh
           DoCmd.Save
           Call MsgBox("Draw: " & Me.DrawName & " has been saved" & venewline & vbNewLine & "The form will CLOSE", vbInformation, "Draw Saved")
           Call CloseMe_Click
        End If
    End If

continue:
End Sub
