Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim FlawTest   As Integer
Dim RST        As DAO.Recordset
Private Sub Form_Current()
    Me.SaveTax.Enabled = False
End Sub

Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.TaxLevel = SplitString(Me.OpenArgs, "|", 0)
        Me.TaxTypeID = SplitString(Me.OpenArgs, "|", 1)
        Me.TaxLevelLookupID = SplitString(Me.OpenArgs, "|", 2)
        Me.BaseTaxRate = Format(SplitString(Me.OpenArgs, "|", 3), "#0.####0%")
        Me.Mode = SplitString(Me.OpenArgs, "|", 4)
    End If
    'If Me.Mode = "Add" Then Me.StartsOnDate = ""
    If Me.Mode = "Edit" Then
        Me.DeleteMe.Enabled = True
        Me.TaxTypeID.Enabled = False
        Call SetRecordSource
    Else
        Me.DeleteMe.Enabled = False
        Me.TaxTypeID.Enabled = True
    End If
    If TaxLevel = "City" And DCount("ID", "TaxCitiesBase", "TaxCityID=" & TaxLevelLookupID) = 0 Then Me.TaxTypeID = 6
    If TaxLevel = "State" And DCount("ID", "TaxStatesBase", "TaxStatesID=" & TaxLevelLookupID) = 0 Then Me.TaxTypeID = 6
    If TaxLevel = "County" And DCount("ID", "TaxCountiesBase", "TaxCountyID=" & TaxLevelLookupID) = 0 Then Me.TaxTypeID = 6
    If Me.TaxTypeID = 6 Then Call SetRecordSource
End Sub
Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub DeleteMe_Click()
    If Me.EditedTaxID & 0 = 0 Then
        Call Dialog.RichBox("You Must Select a TAX before you can delete it", vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
    End If
    Dim ResposeX As VbMsgBoxResultEx
    Dim msg1   As String
    msg1 = msg1 & "<b><Font Size=3>Delete!</font></b><br/>"
    msg1 = msg1 & "You are about to delete a<b> " & Me.TaxLevel & " " & TaxTypeID.Column(1) & "</b><br/>"
    msg1 = msg1 & "  <b>Are you sure?</b><br/>"
    msg1 = msg1 & "  Please select..."
    ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
    If ResposeX = vbBt2 Then Exit Sub
    If ResposeX = vbBt1 Then
        Dim TargetX As String
        TargetX = "TaxStatesBase"
        If Me.TaxLevel = "City" Then TargetX = "TaxCitiesBase"
        If Me.TaxLevel = "County" Then TargetX = "TaxCountiesBase"
        SQLz = "Delete From " & TargetX & " WHERE ID=" & Me.EditedTaxID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Call RefreshTaxForms
        If ActiveTaxStatus(Me.TaxLevel, TaxTypeID, TaxLevelLookupID) Like "*Valid*" Then
            ResposeX = Dialog.RichBox(ActiveTaxStatus(Me.TaxLevel, TaxTypeID, TaxLevelLookupID), vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
        Else
            ResposeX = Dialog.RichBox(ActiveTaxStatus(Me.TaxLevel, TaxTypeID, TaxLevelLookupID), vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
        End If
        DoCmd.close acForm, Me.Form.Name
    End If
End Sub
Private Sub StartsOnDate_AfterUpdate()
    If Me.EndsOnDate <= Me.StartsOnDate Then
        MsgBox "You can't set an End Date on or before the Start Date", vbCritical, AtlaasVersion
        Me.EndsOnDate = ""
    End If
    If IsDate(Me.StartsOnDate) Then Call ValidateNewTax Else Me.ValidateMsg = ""
End Sub
Private Sub StartsOnDate_Change()
    On Error GoTo FixMe
    If Me!StartsOnDate.Text & "test" = "test" Then
        If Me.Dirty Then Me.Dirty = False
        Exit Sub
    End If
    If IsDate(CDate(Me!StartsOnDate.Text)) Then
        If Len(Me!StartsOnDate.Text) < 8 Then Exit Sub
        If Len(Me!StartsOnDate.Text) = 8 And CInt(SplitString(SplitString(Me!StartsOnDate.Text, "/", 1), "/", 0)) > 9 Then Exit Sub
        If CInt(SplitString(Me!StartsOnDate.Text, "/", 0)) > 9 And Len(Me!StartsOnDate.Text) < 10 Then Exit Sub
        If Me.Dirty Then Me.Dirty = False
    End If
FixMe:
    On Error GoTo -1
End Sub
Private Sub EndsOnDate_AfterUpdate()
    If Me.EndsOnDate <= Me.StartsOnDate Then
        MsgBox "You can't set an End Date on or before the Start Date", vbCritical, AtlaasVersion
        Me.EndsOnDate = ""
    End If
    If IsDate(Me.StartsOnDate) Then Call ValidateNewTax
End Sub

Private Sub EndsOnDate_Change()
    On Error GoTo FixMe
    If Me!EndsOnDate.Text & "test" = "test" Then
        If Me.Dirty Then Me.Dirty = False
        Exit Sub
    End If
    If IsDate(CDate(Me!EndsOnDate.Text)) Then
        If Len(Me!EndsOnDate.Text) < 8 Then Exit Sub
        If Len(Me!EndsOnDate.Text) = 8 And CInt(SplitString(SplitString(Me!EndsOnDate.Text, "/", 1), "/", 0)) > 9 Then Exit Sub
        If CInt(SplitString(Me!EndsOnDate.Text, "/", 0)) > 9 And Len(Me!EndsOnDate.Text) < 10 Then Exit Sub
        If Me.Dirty Then Me.Dirty = False
    End If
FixMe:
    On Error GoTo -1
End Sub

Private Sub SetRecordSource()
    If Me.TaxLevel = "State" Then
        Me.EventDescription = Me.Mode & "ing a <b>" & Me.TaxTypeID.Column(1) & "</b> Rate to <b>" & Me.TaxLevel & " </b>of <b>" & DLookup("State", "vw_CCSZ_Lookup", "StateID=" & Me.TaxLevelLookupID) & "</b>"
        SQLz = "SELECT * From vw_Tax_StateBaseQ " & _
               "WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxStatesID=" & TaxLevelLookupID & " " & _
               "ORDER BY StartsOn DESC;"
    End If
    If Me.TaxLevel = "County" Then
        Me.EventDescription = "Adding a <b>" & Me.TaxTypeID.Column(1) & "</b> Rate to <b>" & DLookup("County", "vw_CCSZ_Lookup", "CountyID=" & Me.TaxLevelLookupID) & " County</b>"
        SQLz = "SELECT * From vw_Tax_CountyBaseQ " & _
               "WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxCountyID=" & TaxLevelLookupID & " " & _
               "ORDER BY StartsOn DESC;"
    End If
    If Me.TaxLevel = "City" Then
        Me.EventDescription = "Adding a <b>" & Me.TaxTypeID.Column(1) & "</b> Rate to <b>" & Me.TaxLevel & " </b>of <b>" & DLookup("City", "vw_CCSZ_Lookup", "CityID=" & Me.TaxLevelLookupID) & "</b>"
        SQLz = "SELECT * From vw_Tax_CityBaseQ " & _
               "WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxCityID=" & TaxLevelLookupID & " " & _
               "ORDER BY StartsOn DESC;"
    End If
    Me.Tax_AnySub.Form.RecordSource = SQLz
    Me.ValidateMsg = ""
    Me.StartsOnDate.BackColor = RGB(255, 255, 255)
    If IsDate(Me.StartsOnDate) Then Call ValidateNewTax

End Sub

Public Sub ValidateNewTax()
    Dim PrevStartDate As Date
    Dim PrevEndDate As Date
    Dim NewStartDate As Date
    Dim NewEndDate As Date
    Dim msg1, FiveSpaces, TenSpaces, FourteenSpaces As String
    Dim xi     As Integer
    Me.PreviousTaxID = 0
    Me.ValidateMsg = ""
    If Me.TaxLevel = "State" Then SQLz = "SELECT * From vw_Tax_StateBaseQ WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxStatesID=" & TaxLevelLookupID & " AND IsActive = True AND ID<>" & Me.EditedTaxID & " "
    If Me.TaxLevel = "County" Then SQLz = "SELECT * From vw_Tax_CountyBaseQ WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxCountyID=" & TaxLevelLookupID & " AND IsActive = True AND ID<>" & Me.EditedTaxID & " "
    If Me.TaxLevel = "City" Then SQLz = "SELECT * From vw_Tax_CityBaseQ WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxCityID=" & TaxLevelLookupID & " AND IsActive = True AND ID<>" & Me.EditedTaxID & " "
    SQLz = SQLz & "Order By EndsOn Desc"
    FlawTest = 0
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        FlawTest = 1
        For xi = 1 To RST.RecordCount
            If RST!StartsOn = Me.StartsOnDate Then FlawTest = 2
            If IsDate(Me.EndsOnDate) Then
                If Me.StartsOnDate < RST!StartsOn And Me.EndsOnDate > RST!StartsOn Then FlawTest = 3
            Else
                If Me.StartsOnDate < RST!StartsOn And IsNull(RST!EndsOn) Then FlawTest = 4
            End If
            RST.MoveNext
        Next xi
        RST.MoveFirst
        If FlawTest > 0 Then GoTo DoMsg
    End If
    If RST.RecordCount > 0 Then
        RST.MoveLast
        RST.MoveFirst
        Dim StartsOnx As Date
        Dim EndsOnx As Date
        Dim ii As Integer
        Dim Overlap As Boolean
        For ii = RST.RecordCount To 1 Step -1
            EndsOnx = Nz(Me.EndsOn, "12/31/2050")
            StartsOnx = Me.StartsOn
            If EndsOn <> "12/31/2050" And StartsOnx < RST!StartsOn And EndsOnx < RST!EndsOn Then FlawTest = 5
            RST.MoveNext
        Next ii
    End If
DoMsg:
    Select Case FlawTest
    Case 0
        msg1 = msg1 & "<font color=green><b>Valid!</b></font> This<b> " & Me.TaxTypeID.Column(1) & "</b> rate of <b>" & Format(Me.BaseTaxRate, "##.0###%") & "</b> with a start on date of <b>" & Me.StartsOnDate & "</b>"
        If IsDate(Me.EndsOnDate) Then msg1 = msg1 & ", with an<b> expiration date</b> of<b> " & Me.EndsOnDate & "</b>" Else msg1 = msg1 & ". Currently set to continue in perpetuity"
        If isTaxActive(Me.StartsOn, Nz(Me.EndsOn, "1/1/2050")) = False Then
            msg1 = msg1 & "<br/><b>Note:</b> this tax will <b>not</b> be active!"
        Else
            msg1 = msg1 & ", as an<b> active</b> tax rate"
        End If
        'msg1 = msg1 & "<br/>" & HSpaces(10) & "There were no previous<b> " & Me.TaxTypeId.Column(1) & "</b> rates set for this " & Me.TaxLevel
    Case 1
        msg1 = msg1 & "<font color=green><b>Valid!</b></font> This<b> " & Me.TaxTypeID.Column(1) & "</b> rate of <b>" & Format(Me.BaseTaxRate, "##.0###%") & "</b> with a start on date of  <b>" & Me.StartsOnDate & "</b>"
        If IsDate(Me.EndsOnDate) Then msg1 = msg1 & ", with an<b> expiration date</b> of<b> " & Me.EndsOnDate & "</b>" Else msg1 = msg1 & ". Currently set to continue in perpetuity"
        If isTaxActive(Me.StartsOn, Nz(Me.EndsOn, "1/1/2050")) = False Then
            msg1 = msg1 & "<br/><b>Note:</b> this tax will <b>not</b> be active!"
        Else
            msg1 = msg1 & ", as an<b> active</b> tax rate"
        End If
        If IsDate(Me.StartsOnDate) And Me.StartsOnDate > RST!StartsOn Then msg1 = msg1 & ". This rate replaces the 'previous' " & RST!TaxFullName & " rate of <b>" & Format(RST!BaseTaxRate, "##.0###%") & "</b>, which will expire on <b>" & DateAdd("d", -1, Me.StartsOnDate) & "</b>"
        If IsDate(Me.StartsOnDate) And Me.StartsOnDate > RST!StartsOn Then Me.PreviousTaxID = RST!ID
    Case 2
        msg1 = msg1 & "<font color=red><b>Conflict!</b></font> This " & RST!TaxFullName & " has a <b>Starts On</b> date of <b>" & Me.StartsOnDate & "</b><br/>"
        msg1 = msg1 & HSpaces(14) & "which is the<b><font color=red> same</font></b> as the existing " & RST!TaxFullName & " <b>Starts On</b> date (<i>" & RST!StartsOn & "</i>)"
    Case 3
        msg1 = msg1 & "<font color=red><b>Overlap! </b></font> This " & RST!TaxFullName & " has a <b>Starts On</b> date of <b>" & Me.StartsOnDate & "</b><br/>"
        msg1 = msg1 & HSpaces(15) & "Which preceeds the 'Active'<b> " & RST!TaxFullName & "</b> Start On date <b> " & RST!StartsOn & "</b><br/>"
        msg1 = msg1 & HSpaces(15) & "The<b> End On</b> date selected<b> " & Me.EndsOnDate & "</b> overlaps with the Active " & RST!TaxFullName & "<br />"
        msg1 = msg1 & HSpaces(15) & "The Ends On date must be " & DateAdd("d", -1, RST!StartsOn) & " or earlier"
    Case 4
        msg1 = msg1 & "<font color=red><b>2 Active!</b></font> This " & RST!TaxFullName & " has a <b>Starts On</b> date of <b>" & Me.StartsOnDate & "</b><br/>"
        msg1 = msg1 & HSpaces(15) & " Which preceeds the currently 'Active'<b> " & RST!TaxFullName & "</b> Start On date, set at " & RST!StartsOn & "</b><br/>"
        msg1 = msg1 & HSpaces(15) & " The<b> Ends On date</b> is blank (null) for both this and the currently Active " & RST!TaxFullName & "<br />"
        msg1 = msg1 & HSpaces(15) & " The would create (2) two active " & RST!TaxFullName & " rates for this " & Me.TaxLevel
    Case 5
        msg1 = msg1 & "<font color=red><b>Date Range Overlap!</font></b><br/>"
        msg1 = msg1 & "This " & Me.TaxLevel & " your editing, has a Starts On date that is before another Tax's Ends On Date <br />"
        msg1 = msg1 & "that will cause an<b> Overlap </b>with another " & Me.TaxLevel & " tax<br />"
        msg1 = msg1 & "Please<b> correct </b>the date ranges."
        ' Me.StartsOn = ""
        'Me.EndsOn = ""
    End Select
    If FlawTest < 2 Then
        Me.SaveTax.Enabled = True
        Me.StartsOnDate.BackColor = RGB(255, 255, 255)
        Me.EndsOnDate.BackColor = RGB(255, 255, 255)
        On Error Resume Next
        If RST.RecordCount > 0 And Format(RST!BaseTaxRate, "percent") = Format(Me.BaseTaxRate, "percent") Then _
           msg1 = msg1 & ".<font color=DarkOrange><b> Please Note:</b></font> this tax rate and the previous tax rate are the<b> same</b> (" & Format(RST!BaseTaxRate, "##.0###%") & ")"
        On Error GoTo 0
    End If
    If FlawTest > 2 Then Me.SaveTax.Enabled = False
    If FlawTest > 3 Then Me.EndsOnDate.BackColor = RGB(240, 250, 180)
    Me.ValidateMsg = msg1
    'Me.EndsOnDate.SetFocus
End Sub

Private Sub SaveTax_Click()
    If Me.Dirty Then Me.Dirty = False
    Dim TaxIDLevelx As String
    Dim TaxTableLevelx As String
    SQLz = ""
    TaxIDLevelx = "TaxStatesID"
    TaxTableLevelx = "TaxStatesBase"
    If Me.TaxLevel = "City" Then TaxIDLevelx = "TaxCityID"
    If Me.TaxLevel = "City" Then TaxTableLevelx = "TaxCitiesBase"
    If Me.TaxLevel = "County" Then TaxIDLevelx = "TaxCountyID"
    If Me.TaxLevel = "County" Then TaxTableLevelx = "TaxCountiesBase"

    If Me.Mode = "Edit" Then
        SQLz = SQLz & "UPDATE TaxBaseTemp As S"
        SQLz = SQLz & " INNER JOIN " & TaxTableLevelx & " As T"
        SQLz = SQLz & " ON S.EditedTaxID = T.ID"
        SQLz = SQLz & " SET"
        SQLz = SQLz & " T.TaxTypeId = S.TaxTypeId"
        SQLz = SQLz & ",T.BaseTaxRate = S.BaseTaxRate"
        SQLz = SQLz & ",T.StartsOn = S.StartsOn"
        SQLz = SQLz & ",T.EndsOn = S.EndsOn"
        SQLz = SQLz & ",T.TaxNote = S.TaxNote"
        SQLz = SQLz & ",T.TaxScript = S.TaxScript;"
    Else
        SQLz = SQLz & "INSERT INTO " & TaxTableLevelx & ""
        SQLz = SQLz & "(" & TaxIDLevelx & ""
        SQLz = SQLz & ",TaxTypeId"
        SQLz = SQLz & ",BaseTaxRate"
        SQLz = SQLz & ",StartsOn"
        SQLz = SQLz & ",EndsOn"
        SQLz = SQLz & ",TaxNote"
        SQLz = SQLz & ",TaxScript)"
        SQLz = SQLz & " Select"
        SQLz = SQLz & " TaxLevelLookupID"
        SQLz = SQLz & ",TaxTypeId"
        SQLz = SQLz & ",BaseTaxRate"
        SQLz = SQLz & ",StartsOn"
        SQLz = SQLz & ",EndsOn"
        SQLz = SQLz & ",TaxNote"
        SQLz = SQLz & ",TaxScript"
        SQLz = SQLz & " From TaxBaseTemp"
    End If
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    If Me.PreviousTaxID > 0 Then
        SQLz = ""
        SQLz = SQLz & "Update " & TaxTableLevelx & ""
        SQLz = SQLz & " Set EndsOn='" & DateAdd("d", -1, Me.StartsOnDate) & "' "
        SQLz = SQLz & ",TaxScript = TaxScript & '<b>|NOTE:|</b>rate end = " & DateAdd("d", -1, Me.StartsOnDate) & "'"
        SQLz = SQLz & " WHERE ID=" & Me.PreviousTaxID & ";"

        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Call RefreshTaxForms
    Call Dialog.RichBox("<font size=2>This<b> " & Me.TaxLevel & "'s " & Me.TaxTypeID.Column(1) & "</font></b><br />has been saved, the from will now close", vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
    DoCmd.close acForm, Me.Form.Name
End Sub


Private Sub BaseTaxRate_AfterUpdate()
    If Me.BaseTaxRate < 0 Then
        MsgBox "You Cannot Have A Tax Rate Less Than Zero", vbCritical, AtlaasVersion
        Me.BaseTaxRate = 0
        Me.ValidateMsg = "<font color=red><b>Error!</b></font><Br />You cannot have a Tax Rate less than<b> Zero</b>"
        Me.SaveTax.Enabled = False
        Exit Sub
    End If
    If IsDate(Me.StartsOnDate) Then Call ValidateNewTax
End Sub

Private Sub TaxTypeID_Change()
    Call SetRecordSource
End Sub
