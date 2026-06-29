Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim FlawTest   As Integer
Dim RST        As DAO.Recordset
Private ranCurrent As Boolean


Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.TaxLevel = SplitString(Me.OpenArgs, "|", 0)
        Me.TaxTypeID = SplitString(Me.OpenArgs, "|", 1)
        Me.TaxLevelLookupID = SplitString(Me.OpenArgs, "|", 2)
        Me.CurrTaxRate = CDbl(SplitString(Me.OpenArgs, "|", 3))
        Me.Mode = SplitString(Me.OpenArgs, "|", 4)
    Else
        Me.TaxLevel = "State"
        Me.TaxLevelLookupID = 11
        Me.TaxTypeID = 1
    End If
    Call SetRecordSource
    DoCmd.SetWarnings False
    If Me.Mode = "Add" Then DoCmd.RunSQL ("Delete From TaxExceptionsNAICSTemp")
    If Me.Mode = "Add" Then Me.DeleteMe.Enabled = False Else Me.DeleteMe.Enabled = True
    Me.NAICSList.Requery
    DoCmd.SetWarnings True
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax Else Call SetColors
End Sub
Private Sub DeleteMe_Click()
If Me.ID & 0 = 0 Then
Call Dialog.RichBox("You Must Select a TAX before you can delete it", vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
End If
    Dim ResposeX As VbMsgBoxResultEx
    Dim msg1   As String
    msg1 = msg1 & "<b><Font Size=3>Delete!</font></b><br/>"
    msg1 = msg1 & "You are about to delete the exception:<br/><b> " & Me.ExceptionDescription & "</b><br/>"
    msg1 = msg1 & "  <b>Are you sure?</b><br/>"
    msg1 = msg1 & "  Please select..."
    ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
    If ResposeX = vbBt2 Then Exit Sub
    If ResposeX = vbBt1 Then
        Dim TargetX As String
        TargetX = "TaxStatesExceptions"
        If Me.TaxLevel = "City" Then TargetX = "TaxCitiesExceptions"
        If Me.TaxLevel = "County" Then TargetX = "TaxCountiesExceptions"
        SQLz = "Delete From " & TargetX & " WHERE ID=" & Me.ID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Call RefreshTaxForms
        DoCmd.close acForm, Me.Form.Name
    End If
End Sub

Private Sub ExceptionDescription_AfterUpdate()
 If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax Else Call SetColors
End Sub

Private Sub AddNAICS_Click()
    If Me.NAICSCbo & "test" <> "test" Then
        If DCount("Indx", "TaxExceptionsNAICSTemp", "NAICS='" & Me.NAICSCbo & "'") > 0 Then
            MsgBox Me.NAICSDesc & " is already in the list", vbCritical, AtlaasVersion
            Exit Sub
        End If
        SQLz = "Insert into TaxExceptionsNAICSTemp (NAICS, NAICSDesc, TaxLevel) Values('" & Me.NAICSCbo & "','" & Me.NAICSDesc & "','" & Me.TaxLevel & "')"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.NAICSList.Requery
    End If
If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub DeleteNAICS_Click()
    SQLz = "Delete * From TaxExceptionsNAICSTemp WHERE Indx =" & Me.NAICSList.Column(0) & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.NAICSList.Requery
If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub


Private Sub StartsOnDate_Change()
    On Error GoTo FixMe
    If Not IsDate(CDate(Me!StartsOnDate.Text)) Then Exit Sub
    If IsDate(CDate(Me!StartsOnDate.Text)) Then
        If Len(Me!StartsOnDate.Text) < 8 Then Exit Sub
        If Len(Me!StartsOnDate.Text) = 8 And CInt(SplitString(SplitString(Me!StartsOnDate.Text, "/", 1), "/", 0)) > 9 Then Exit Sub
        If CInt(SplitString(Me!StartsOnDate.Text, "/", 0)) > 9 And Len(Me!StartsOnDate.Text) < 10 Then Exit Sub
        Me.Dirty = False
    End If
FixMe:
End Sub
Private Sub StartsOnDate_AfterUpdate()
Me.TaxScript = ""
  If CompareDates(Me.StartsOnDate.Name) = True Then
  Call ValidateNewTax
  Else
  Me!StartsOnDate.SetFocus
 ' Me!StartsOnDate.Undo = ""
  End If
 ' If CompareDates("StartsOnDate") = True Then Call ValidateNewTax Else Me.StartsOnDate.SetFocus
End Sub
Private Sub StartsOnDate_LostFocus()
If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax Else Me!StartsOnDate.SetFocus
End Sub
Private Sub EndsOnDate_Change()
    On Error GoTo FixMe
    If Not IsDate(CDate(Me!EndsOnDate.Text)) Then Exit Sub
    If IsDate(CDate(Me!EndsOnDate.Text)) Then
        If Len(Me!EndsOnDate.Text) < 8 Then Exit Sub
        If Len(Me!EndsOnDate.Text) = 8 And CInt(SplitString(SplitString(Me!EndsOnDate.Text, "/", 1), "/", 0)) > 9 Then Exit Sub
        If CInt(SplitString(Me!EndsOnDate.Text, "/", 0)) > 9 And Len(Me!EndsOnDate.Text) < 10 Then Exit Sub
        Me.Dirty = False
    End If
FixMe:
    On Error GoTo -1
End Sub
Private Sub EndsOnDate_AfterUpdate()
    Me.TaxScript = ""
    If CompareDates(Me.EndsOnDate.Name) = True Then Call ValidateNewTax Else Me!EndsOnDate.SetFocus
End Sub
Private Sub EndsOnDate_LostFocus()
If CompareDates(Me.EndsOnDate.Name) = True Then Call ValidateNewTax Else Me!EndsOnDate.SetFocus
End Sub
Public Function CompareDates(WhichDate As String) As Boolean
    Dim StartD As Date
    Dim EndD   As Date
    Dim ErrorMsg As String
    CompareDates = False
    If IsNull(Len(Me(WhichDate))) Or Len(Me(WhichDate)) < 8 Then Exit Function
    If Not IsDate(Me.StartsOnDate) And IsDate(Me.EndsOnDate) Then ErrorMsg = ErrorMsg & "You must set a Start Date" & vbNewLine
    If IsDate(Me.StartsOnDate) Then StartD = CDate(Me.StartsOnDate)
    If Len(Me(WhichDate)) < 7 Then ErrorMsg = "Please input a Proper Date (" & WhichDate & ")" & vbNewLine
    If Len(Me(WhichDate)) > 6 And Not IsDate(Me(WhichDate)) Then ErrorMsg = "Please input a Proper Date (" & WhichDate & ")" & vbNewLine
    If IsDate(Me.EndsOnDate) Then
        EndD = CDate(Me.EndsOnDate)
    Else
        If IsDate(Me.StartsOnDate) Then EndD = DateAdd("d", 1, CDate(Me.StartsOnDate))
    End If
    If EndD <= StartD Then ErrorMsg = ErrorMsg & "You cannot set an End Date on or before a Start Date" & vbNewLine
    If EndD < #1/1/1900# Then ErrorMsg = ErrorMsg & "Please input a valid End Date (i.e.: '1/1/1900' or greater)" & vbNewLine
    If StartD < #1/1/1900# Then ErrorMsg = ErrorMsg & "Please input a valid Start Date (i.e.: '1/1/1900' or greater)" & vbNewLine
    If IsDate(Me.StartsOnDate) And Not IsDate(Me.EndsOnDate) And ErrorMsg = "" Then
        CompareDates = True
        Exit Function
    End If
    If ErrorMsg <> "" Then
        MsgBox "Error!" & vbNewLine & Trim(ErrorMsg) & "Please correct", vbCritical, AtlaasVersion
        CompareDates = False
        Me.SaveTax.Enabled = False
        Me.TaxScript = ""
       Me.TaxScript = "<font color=red><b>" & ErrorMsg & "!</Font></b>"
      '  If WhichDate = "StartsOnDate" Then Me.StartsOnDate.SetFocus Else Me.EndsOnDate.SetFocus
      '  If WhichDate = "StartsOnDate" Then Me.StartsOnDate = "" Else Me.EndsOnDate.Text = ""
    Else
    Call ValidateNewTax
    End If
End Function

Private Sub ExceptionDescription_Click()
Call ValidateNewTax
End Sub



Private Sub Form_Current()
    Me.SaveTax.Enabled = False
     If ranCurrent = True Then Exit Sub     ' prevents double execution
    ranCurrent = True
End Sub

Public Sub SetRecordSource()
Dim TaxTypeX As String
TaxTypeX = DLookup("TaxAbbreviation", "TaxType", "ID=" & Me.TaxTypeID)
    If Me.TaxLevel = "State" Then
       Me.TaxStatesID = Me.TaxLevelLookupID
        Me.EventDescription = Me.Mode & "ing an 'Exception' to the <b>" & TaxTypeX & "</b> Rate, <b>" & Me.TaxLevel & " </b>of <b>" & DLookup("StateName", "TaxStates", "ID=" & Me.TaxLevelLookupID) & "</b>"
        SQLz = "SELECT * From vw_TaxStatesExceptionsQ " & _
               "WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxStatesID=" & TaxLevelLookupID & " " & _
               "ORDER BY StartsOn DESC;"
    End If
    If Me.TaxLevel = "County" Then
    Me.TaxCountiesID = Me.TaxLevelLookupID
        Me.EventDescription = Me.Mode & "ing an 'Exception' to the <b>" & TaxTypeX & "</b> Rate, <b>" & DLookup("CountyName", "TaxCounties", "ID=" & Me.TaxLevelLookupID) & " County</b>"
        SQLz = "SELECT * From vw_TaxCountiesExceptionsQ " & _
               "WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxCountiesID=" & TaxLevelLookupID & " " & _
               "ORDER BY StartsOn DESC;"
    End If
    If Me.TaxLevel = "City" Then
        Me.TaxCitiesID = Me.TaxLevelLookupID
        Me.EventDescription = Me.Mode & "ing an 'Exception' to the <b>" & TaxTypeX & "</b> Rate, <b>" & Me.TaxLevel & " </b>of <b>" & DLookup("CityName", "TaxCities", "ID=" & Me.TaxLevelLookupID) & "</b>"
        SQLz = "SELECT * From vw_TaxCitiesExceptionsQ " & _
               "WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxCitiesID=" & TaxLevelLookupID & " " & _
               "ORDER BY StartsOn DESC;"
    End If
    Me.Tax_Any_ExceptionSub.Form.RecordSource = SQLz
    'Me.ValidateMsg = ""
    Me.StartsOnDate.BackColor = RGB(255, 255, 255)
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub

Public Sub ValidateNewTax()
  '  Me.ValidateMsg = ""
    Call SetColors
    Dim msg1   As String
    Dim Script As String
        If Me.TaxRule = True Then
        Script = "<font color=Navy>This Tax<font color=DarkOrange><b> Replaces</b></font> the Base Tax " & Script & "</font> "
        Me.EffectiveTax = CDbl(Nz(Me.TaxRate, 0))
    Else
        Me.EffectiveTax = CDbl(Nz(Me.[CurrTaxRate], 0) + Nz(Me.[TaxRate], 0))
    End If
    If Nz(Me.EffectiveTax, 0) < 0 Then msg1 = msg1 & "<br />You<b> Cannot </b>have an 'Effective Tax Rate' that is less than Zero"
    If Nz(Me.ExceptionDescription, "") = "" Then msg1 = msg1 & "<br />You<b> Must </b>Enter an <b>Exception</b> before you can start"
    If IsNull(Me.TaxRate) Then msg1 = msg1 & "<br />You <b> Must </b> set an <b> Tax Rate </b> even if <b>0%</b>, before you can start"
    If msg1 <> "" Then
        Beep
        msg1 = "<font color=red><b>Error! </b></font> " & msg1
        Me.TaxScript = msg1
        Me.SaveTax.Enabled = False
        Exit Sub
    End If
    msg1 = ""
    Script = ""
    If Me.AppliesToTrailers = True Then Script = Script & "<b> Trailers</b>"
    If Me.AppliesToTractors = True Then
        If Script <> "" Then Script = Script & " - "
        Script = Script & "<b> Tractors</b>"
    End If

    If Me.AppliesToStraightTrucks = True Then
        If Script <> "" Then Script = Script & " - "
        Script = Script & "<b> Straight Trucks</b>"
    End If

    If IsNumeric(Me.TermExceedsMonths) Then Script = Script & " with a<b> Lease Term</b> of more than<b> " & Me.TermExceedsMonths & " </b>months"
    If IsNumeric(Me.WeightLimitMinLbs) And IsNumeric(Me.WeightLimitMaxLbs) Then
        Script = Script & ", including units that weight between<b> " & Format(Me.WeightLimitMinLbs, "#,#") & " </b>lbs and<b> " & Format(Me.WeightLimitMaxLbs, "#,#") & " </b>lbs"
    Else
        If IsNumeric(Me.WeightLimitMinLbs) Then Script = Script & ", including units that weight a least<b> " & Format(Me.WeightLimitMinLbs, "#,#") & " </b>lbs"
        If IsNumeric(Me.WeightLimitMaxLbs) Then Script = Script & ", including units that weight less than <b> " & Format(Me.WeightLimitMaxLbs, "#,#") & " </b>lbs"
    End If

    If DCount("Indx", "TaxExceptionsNAICSTemp") > 0 Then
        Set RST = CurrentDb.OpenRecordset("SELECT * FROM TaxExceptionsNAICSTemp Order By NAICS", dbOpenSnapshot)
        On Error Resume Next
        RST.MoveLast
        RST.MoveFirst
        On Error GoTo 0
        If RST.RecordCount > 1 Then Script = Script & ", for clients with NAICS Codes:" Else Script = Script & ", for clients with the NAICS Code:"
        Me.IncludedNAICS = ""
        For ii = 1 To RST.RecordCount
            If ii = 1 Then Me.IncludedNAICS = RST!NAICS
            If ii > 1 Then Me.IncludedNAICS = Me.IncludedNAICS & "|" & RST!NAICS
            Script = Script & "<b> " & RST!NAICS & "</b>"
            If ii >= RST.RecordCount Then Script = Script & "." Else Script = Script & ","
            RST.MoveNext
        Next
        RST.close
        Set RST = Nothing
    End If
    If IsNumeric(Me.TaxMaxValue) Then Script = Script & " This Exception will not exceed<b> " & Format(Me.TaxMaxValue, "Currency") & "</b>"
    If Me.TaxRule = True Then
        Script = "</b><font color=Navy>This Tax<font color=DarkOrange><b> Replaces</b></font> the Base Tax" & Script & "</font> "
        Me.EffectiveTax = CDbl(Nz(Me.TaxRate, 0))
    Else
        Me.EffectiveTax = CDbl(Nz(Me.[CurrTaxRate], 0) + Nz(Me.[TaxRate], 0))
    End If
    If Script <> "" And Me.TaxRule = False Then
        If Me.TaxRate >= 0 Then Script = "<font color=Navy>This Tax<font color=Red><b> Increase </b></font>applies to" & Script & "</font>"
        If Me.TaxRate <= 0 Then Script = "<font color=Navy>This Tax<font color=Green><b> Reduction </b></font>applies to" & Script & "</font>"
    End If
    If Script <> "" Then Me.TaxScript = Script Else Me.TaxScript = ""
    Dim PrevStartDate As Date
    Dim PrevEndDate As Date
    Dim NewStartDate As Date
    Dim NewEndDate As Date
    Dim ExcDesc As String
    Dim xi     As Integer
    If Me.TaxLevel = "State" Then SQLz = "SELECT * From vw_TaxStatesExceptionsQ WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxStatesID=" & TaxLevelLookupID & " AND ExceptionDescription = '" & Me.ExceptionDescription & "' AND IsActive = True AND ID <>" & Me.ID & " "
    If Me.TaxLevel = "County" Then SQLz = "SELECT * From vw_TaxCountiesExceptionsQ WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxCountiesID=" & TaxLevelLookupID & " AND ExceptionDescription = '" & Me.ExceptionDescription & "' AND IsActive = True AND ID <>" & Me.ID & " "
    If Me.TaxLevel = "City" Then SQLz = "SELECT * From vw_TaxCitiesExceptionsQ WHERE TaxTypeID =" & Me.TaxTypeID & " AND TaxCitiesID=" & TaxLevelLookupID & " AND ExceptionDescription = '" & Me.ExceptionDescription & "' AND IsActive = True AND ID <>" & Me.ID & " "
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

            ' If Nz(RST!EndsOn, CDate("1/1/1900")) >= Me.StartsOnDate Then FlawTest = 3
            RST.MoveNext
        Next xi
        RST.MoveFirst
        'Me.ExceptionDescription = Nz(RST!ExceptionDescription, "")
    End If
    ExcDesc = Me.ExceptionDescription
    Select Case FlawTest
    Case 0
        msg1 = msg1 & "<font color=green><b>Valid!</b></font> This<b> " & ExcDesc & "</b> rate of <b>" & Format(Me.TaxRate, "##.0###%") & "</b> will start on <b>" & Me.StartsOnDate & "</b><br />" & HSpaces(10)
        If IsDate(Me.EndsOnDate) Then msg1 = msg1 & "and will<b> expire</b> on<b> " & Me.EndsOnDate & "</b>" Else msg1 = msg1 & "That tax rate will continue in perpetuity"
        If IsDate(Me.EndsOnDate) And Me.EndsOnDate <= Date Then msg1 = msg1 & "<font color=DarkOrange><b> Please Note:</b></font> this exception will <b>not</b> be active!"
        If IsDate(Me.EndsOnDate) And Me.EndsOnDate > Date Then msg1 = msg1 & ". This will be an<b> active</b> exception"
    '    If IsDate(Me.EndsOnDate) And Me.EndsOnDate <= Date Then msg1 = msg1 & " an<b> inactive</b> exception"

      '  msg1 = msg1 & "<br/>" & HSpaces(10) & "There were no previous<b> " & ExcDesc & "</b> rates set for this " & Me.TaxLevel
    Case 1
        msg1 = msg1 & "<font color=green><b>Valid!</b></font> This<b> " & ExcDesc & "</b> rate of <b>" & Format(Me.TaxRate, "##.0###%") & "</b> will start on <b>" & Me.StartsOnDate & "</b><br />" & HSpaces(10)
        If IsDate(Me.EndsOnDate) Then msg1 = msg1 & "This tax rate will<b> expire</b> on<b> " & Me.EndsOnDate & "</b>" Else msg1 = msg1 & "That tax rate will continue in perpetuity"
        If IsDate(Me.EndsOnDate) And Me.EndsOnDate <= Date Then msg1 = msg1 & "<br/>" & HSpaces(10) & "<b>Note:</b> this tax will <b>not</b> be active!"
        If Not IsDate(Me.EndsOnDate) Or CDate(Nz(Me.EndsOnDate, "1/1/1800")) > Date Then msg1 = msg1 & " This will be an<b> active</b> tax rate"
        If IsDate(Me.StartsOnDate) And Me.StartsOnDate > RST!StartsOn Then msg1 = msg1 & "<br/>" & HSpaces(10) & "This will replace the 'previous'<b> " & RST!ExceptionDescription & "</b> rate of <b>" & Format(RST!TaxRate, "##.0###%") & "</b>,<br/>" & HSpaces(10) & "which will expire on<b> " & DateAdd("d", -1, Me.StartsOnDate) & "</b>"
    Case 2
        msg1 = msg1 & "<font color=red><b>Conflict!</b></font> This " & ExcDesc & " has a <b>Starts On</b> date of <b>" & Me.StartsOnDate & "</b><br/>"
        msg1 = msg1 & HSpaces(14) & "which is the<b><font color=red> same</font></b> as the existing " & RST!ExceptionDescription & " <b>Starts On</b> date (<i>" & RST!StartsOn & "</i>)"
    Case 3
        msg1 = msg1 & "<font color=red><b>Overlap! </b></font> This " & ExcDesc & " has a <b>Starts On</b> date of <b>" & Me.StartsOnDate & "</b><br/>"
        msg1 = msg1 & HSpaces(15) & "Which preceeds the 'Active'<b> " & RST!TaxFullName & "</b> Start On date <b> " & RST!StartsOn & "</b><br/>"
        msg1 = msg1 & HSpaces(15) & "The<b> End On</b> date selected<b> " & Me.EndsOnDate & "</b> overlaps with the Active " & RST!ExceptionDescription & "<br />"
        msg1 = msg1 & HSpaces(15) & "The Ends On date must be " & DateAdd("d", -1, RST!StartsOn) & " or earlier"
    Case 4
        msg1 = msg1 & "<font color=red><b>Two Active!</b></font> This 'new'<b> " & ExcDesc & " </b>has a <b>Starts On</b> date of <b>" & Me.StartsOnDate & "</b><br/>"
        msg1 = msg1 & HSpaces(15) & " Which <i>preceeds</i> the currently 'Active' " & ExcDesc & "'s Start On date (" & RST!StartsOn & ")</b>.<br />"
        msg1 = msg1 & HSpaces(15) & " The<b> Ends On </b>dates are blank for both exceptions<br />"
        msg1 = msg1 & HSpaces(15) & " The would create (2) two active " & ExcDesc & " rates for this " & Me.TaxLevel

    End Select
    If FlawTest < 2 Then
        Me.SaveTax.Enabled = True
        Me.StartsOnDate.BackColor = RGB(255, 255, 255)
        Me.EndsOnDate.BackColor = RGB(255, 255, 255)
        On Error Resume Next
        If RST.RecordCount > 0 And Format(RST!TaxRate, "percent") = Format(Me.TaxRate, "percent") Then _
           msg1 = msg1 & "<font color=DarkOrange><b>Please Note:</b></font> this Exception rate and the Base Tax rate are the<b> same</b> (" & Format(RST!TaxRate, "percent") & ")"
        On Error GoTo 0
    End If
    If FlawTest > 2 Then Me.SaveTax.Enabled = False
    If FlawTest > 3 Then Me.EndsOnDate.BackColor = RGB(240, 250, 180)
    Me.TaxScript = msg1 & ".<br/>" & Script
End Sub

Private Sub NAICSCbo_Click()
 Me.NAICSDesc = Me.NAICSCbo.Column(1)
 Me.AddNAICS.SetFocus
End Sub

Private Sub SaveTax_Click()
    If Me.Form.Dirty Then Me.Form.Dirty = False
    If Me.Mode = "Edit" Then
        SQLz = ""
        SQLz = SQLz & "UPDATE TaxExceptionsTemp AS T INNER JOIN TaxStatesExceptions AS E ON E.ID = T.TaxExceptionID SET"
        SQLz = SQLz & " E.TaxStatesID= T.TaxStatesID"
        SQLz = SQLz & ",E.TaxTypeID= T.TaxTypeID"
        SQLz = SQLz & ",E.TaxExceptionTypeID= T.TaxExceptionTypeID"
        SQLz = SQLz & ",E.ExceptionDescription= T.ExceptionDescription"
        SQLz = SQLz & ",E.TaxRule= T.TaxRule"
        SQLz = SQLz & ",E.TaxRate= T.TaxRate"
        SQLz = SQLz & ",E.FormURL1= T.FormURL1"
        SQLz = SQLz & ",E.FormURL2= T.FormURL2"
        SQLz = SQLz & ",E.IncludedNAICS= T.IncludedNAICS"
        SQLz = SQLz & ",E.AppliesToTractors= T.AppliesToTractors"
        SQLz = SQLz & ",E.AppliesToTrailers= T.AppliesToTrailers"
        SQLz = SQLz & ",E.AppliesToStraightTrucks= T.AppliesToStraightTrucks"
        SQLz = SQLz & ",E.WeightLimitMinLbs= T.WeightLimitMinLbs"
        SQLz = SQLz & ",E.WeightLimitMaxLbs= T.WeightLimitMaxLbs"
        SQLz = SQLz & ",E.TermExceedsMonths= T.TermExceedsMonths"
        SQLz = SQLz & ",E.StartsOn= T.StartsOn"
        SQLz = SQLz & ",E.EndsOn= T.EndsOn"
        SQLz = SQLz & ",E.TaxMaxValue= T.TaxMaxValue"
        SQLz = SQLz & ",E.TaxScript= T.TaxScript"
        SQLz = SQLz & ";"
    End If
    If Me.Mode = "Add" Then
        SQLz = "INSERT INTO TaxStatesExceptions "
        SQLz = SQLz & "(TaxTypeID"
        SQLz = SQLz & ",TaxStatesID"
        SQLz = SQLz & ",TaxExceptionTypeID"
        SQLz = SQLz & ",ExceptionDescription"
        SQLz = SQLz & ",TaxRule"
        SQLz = SQLz & ",TaxRate"
        SQLz = SQLz & ",FormURL1"
        SQLz = SQLz & ",FormURL2"
        SQLz = SQLz & ",IncludedNAICS"
        SQLz = SQLz & ",AppliesToTractors"
        SQLz = SQLz & ",AppliesToTrailers"
        SQLz = SQLz & ",AppliesToStraightTrucks"
        SQLz = SQLz & ",WeightLimitMinLbs"
        SQLz = SQLz & ",WeightLimitMaxLbs"
        SQLz = SQLz & ",TermExceedsMonths"
        SQLz = SQLz & ",StartsOn"
        SQLz = SQLz & ",EndsOn"
        SQLz = SQLz & ",TaxMaxValue"
        SQLz = SQLz & ",TaxScript"
        SQLz = SQLz & ") SELECT"
        SQLz = SQLz & " TaxTypeID"
        SQLz = SQLz & ",TaxStatesID"
        SQLz = SQLz & ",TaxExceptionTypeID"
        SQLz = SQLz & ",ExceptionDescription"
        SQLz = SQLz & ",TaxRule"
        SQLz = SQLz & ",TaxRate"
        SQLz = SQLz & ",FormURL1"
        SQLz = SQLz & ",FormURL2"
        SQLz = SQLz & ",IncludedNAICS"
        SQLz = SQLz & ",AppliesToTractors"
        SQLz = SQLz & ",AppliesToTrailers"
        SQLz = SQLz & ",AppliesToStraightTrucks"
        SQLz = SQLz & ",WeightLimitMinLbs"
        SQLz = SQLz & ",WeightLimitMaxLbs"
        SQLz = SQLz & ",TermExceedsMonths"
        SQLz = SQLz & ",StartsOn"
        SQLz = SQLz & ",EndsOn"
        SQLz = SQLz & ",TaxMaxValue"
        SQLz = SQLz & ",TaxScript"
        SQLz = SQLz & " From TaxExceptionsTemp"
    End If
    If Me.TaxLevel = "City" Then SQLz = Replace(SQLz, "States", "Cities", 1)
    If Me.TaxLevel = "County" Then SQLz = Replace(SQLz, "States", "Counties", 1)
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call RefreshTaxForms
    Call Dialog.RichBox("<font size=2><b>" & Me.ExceptionDescription & "</font></b><br />has been saved<br />...the form will now close", vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
'    MsgBox "Tax Exception was SAVED " & vbNewLine & "The Form Will Close", vbInformation, AtlaasVersion
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub TaxRule_Click()
If Me.TaxRate < 0 Then
MsgBox "You cannot replace a Base Tax rate with a" & vbNewLine & "Negative Exception Rate." & vbNewLine & "Please correct the Exception Rate.", vbCritical, AtlaasVersion
Me.TaxRule = False
End If
Call ValidateNewTax
End Sub

Private Sub TaxRule2_Click()
Call TaxRule_Click
End Sub

Private Sub TaxTypeID_Change()
    Call SetRecordSource
End Sub
Private Sub AppliesToTractors_Click()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub AppliesToTractors2_Click()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub AppliesToStraightTrucks_Click()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub AppliesToStraightTrucks2_Click()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub AppliesToTrailers_Click()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub AppliesToTrailers2_Click()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub ForHireOnly_Click()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub TermExceedsMonths_AfterUpdate()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub TaxMaxValue_AfterUpdate()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub WeightLimitMaxLbs_AfterUpdate()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub


Private Sub WeightLimitMinLbs_AfterUpdate()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub
Private Sub TaxRate_AfterUpdate()
    If CompareDates(Me.StartsOnDate.Name) = True Then Call ValidateNewTax
End Sub

Public Sub SetColors()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Selected" Then
            If Me(Ccrl.Name) > 0 Then Me(Ccrl.Name & "Label").BackColor = RGB(23, 54, 93) Else Me(Ccrl.Name & "Label").BackColor = RGB(216, 216, 216)
            If Me(Ccrl.Name) > 0 Then Me(Ccrl.Name & "Label").ForeColor = RGB(255, 255, 255) Else Me(Ccrl.Name & "Label").ForeColor = RGB(114, 114, 114)
        End If
    Next Ccrl
End Sub
