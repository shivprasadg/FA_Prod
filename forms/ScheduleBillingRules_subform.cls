Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, msg, BillScriptX, PDScriptX As String

Private Sub Form_AfterInsert()
If IsNull(Me.clientID) Then
Call RunScript
Call SaveUpdate_Click
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    Call EnableDisableOptionse("On", "RentFee")
End Sub
Private Sub Form_Load()
    SQLz = "SELECT SGrpID, UnitGroup As Groups, Count(AssetID) AS Units, SchID " & _
           "From vw_SixKeys " & _
           "GROUP BY SGrpID, UnitGroup, SchID " & _
           "HAVING SchID=" & Me.Parent.SchID & " Order BY UnitGroup; "
    Me.GroupList.RowSource = SQLz
End Sub
Private Sub Form_Current()
If Me.Parent.MLOrig <> "FA" Then Exit Sub
    If Me.StartRentOn = 3 Then Me.CustomStartDate.Visible = True Else Me.CustomStartDate.Visible = False
    If Me.Form.NewRecord = True Then

        Dim ItemIndex As Variant
        Dim x, i As Integer
        Dim ScriptZ As String
        x = (GroupList.ListCount - 1)
        If x < 0 Then Exit Sub
        Dim ClientIDx, ScheduleIDx As Integer
        ReDim UnitGroupIdX(1 To x) As Integer
            ClientIDx = Me.Parent.clientID
            ScheduleIDx = Me.Parent.SchID
        For i = 1 To x
            UnitGroupIdX(i) = GroupList.ItemData(i)
        Next i
        GroupList.Selected(1) = True
        Me.UnitGroupId = UnitGroupIdX(1)
        Me.PerDiemType = 1
        Call PerDiemType_Click
        Me.BillType = 1
        Call BillType_Click
        Call RunScript
        If x > 1 Then
            ScriptZ = Me.BillingScript
            For i = 2 To x
                SQLz = "INSERT INTO BillingRulesClient ( ClientId, ScheduleId, UnitGroupId, PerDiemType, PerDiemStartOn, PerDiemFeeToUse, PerDiemFeeValue, " & _
                       "PerDiemFeeTaxable, PerDiemFeeLabel, PerDiemShowFeeColumn, PerDiemRefundFee, BillType, SplitFee, RemitTo, ShowContract, FeeToUse, FeeValue, " & _
                       "FeeTaxable, FeeTerm, FeeLabel, ShowFeeColumn, StartRentOn, RefundFee, DayOfMonth, CustomStartDate, Script ) " & _
                       "SELECT " & ClientIDx & "," & ScheduleIDx & ", " & UnitGroupIdX(i) & ", " & Me.PerDiemType & ", " & Me.PerDiemStartOn & ", " & Me.PerDiemFeeToUse & ", " & Me.PerDiemFeeValue & ", " & _
                       "" & Nz(Me.PerDiemFeeTaxable, 0) & ", " & Nz(Me.PerDiemFeeLabel, 0) & ", " & Me.PerDiemShowFeeColumn & ", " & Me.PerDiemRefundFee & ", " & Me.BillType & ", " & _
                       "" & Me.SplitFee & ", " & Me.RemitTo & ", " & Me.ShowContract & ", " & Me.FeeToUse & ", " & Me.FeeValue & ", " & _
                       "" & Nz(Me.FeeTaxable, 0) & ", " & Me.FeeTerm & ", " & Me.FeeLabel & ", " & Me.ShowFeeColumn & ", " & Me.StartRentOn & ", " & _
                       "" & Me.RefundFee & ", " & Me.DayOfMonth & ", #" & Nz(Me.CustomStartDate, "1/1/1900") & "#, '" & Replace(Nz(ScriptZ, "NO"), "'", """") & "';"

                DoCmd.SetWarnings False
                DoCmd.RunSQL (SQLz)
                DoCmd.SetWarnings True
                'Me.GroupList
            Next i
        End If
        If IsNull(Me.clientID) Then Me.clientID = ClientIDx
    Else
    Call RunScript
    End If
    Me.PerdiemFee.SetFocus

End Sub

'Private Sub BankRent_Click()
'    If Me.BankRent.Column(1) Like "*No Fee*" Then Me.FeeHideCombine = 0
'End Sub

'Private Sub FeeOnly_Click()
'    If Me.FeeOnly = True Then Call ClearRent Else Call EnableRent
'End Sub

Private Sub ClearRent()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Rent" Then
            Me(Ccrl.Name).Value = ""
            Me(Ccrl.Name).Enabled = False
        End If
    Next Ccrl
 '   Me.FeeHideCombine = -1
End Sub
Private Sub EnableRent()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Rent" Then
            Me(Ccrl.Name).Value = ""
            Me(Ccrl.Name).Enabled = True
        End If
    Next Ccrl
  '  Me.FeeHideCombine = -1
End Sub
Private Sub EnableDisableOptionse(OnOff As String, TagX As String)
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = TagX Then
            'Me(cCrl.Name).Value = ""
            If OnOff = "Off" Then Me(Ccrl.Name).Enabled = False Else Me(Ccrl.Name).Enabled = True
        End If
    Next Ccrl
End Sub

Private Sub DayOfMonth_AfterUpdate()
    If Me.DayOfMonth < 1 Or Me.DayOfMonth > 28 Then Me.DayOfMonth = 25
    Call RunScript
End Sub

Private Sub FeeLabel_Click()
    Call RunScript
End Sub

Private Sub FeeTerm_AfterUpdate()
    Call RunScript
End Sub



Private Sub GroupList_DblClick(Cancel As Integer)
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , "SGrpId=" & Me.GroupList & ""
End Sub

Private Sub InvoiceFee_AfterUpdate()
    Call RunScript
End Sub

Private Sub PerDiemFee_AfterUpdate()
    Call RunScript
End Sub

Private Sub PerDiemFeeLabel_Click()
    If Me.PerDiemFeeLabel = 0 Then Me.PerDiemFeeToUse = 0
    Call PerDiemFeeToUse_Click
End Sub

Private Sub PerDiemFeeTaxable_Click()
    Call RunScript
End Sub

Private Sub PerDiemRefundFee_Click()
    Call RunScript
End Sub

Private Sub PerDiemShowFeeColumn_Click()
    Call RunScript
End Sub

Private Sub PerDiemStartOn_Click()
    Call RunScript
End Sub

Private Sub PerDiemType_Click()
    Me.PerDiemStartOn = 2
    Me.PerDiemFeeToUse = 3
    Me.PerDiemFeeLabel = 3
    Me.PerdiemFee = 65
    Me.PerDiemShowFeeColumn = 1
    Me.PerDiemFeeTaxable = 0
    Me.PerDiemRefundFee = 0
    Call EnableDisableOptionse("On", "PDFee")
    Call RunScript
End Sub
Private Sub FeeTaxable_Click()
    Call RunScript
End Sub
Private Sub RefundFee_Click()
    Call RunScript
End Sub

Private Sub SampleInvoice_Click()
DoCmd.OpenReport "BillingRules_Invoice", acViewReport

End Sub

Private Sub SaveUpdate_Click()
'    Me.clientID = Me.Parent.clientID
'    Me.ScheduleId = Me.Parent.SchID
'ReDim S(1 To Selection.Rows.count)
    Dim ItemIndex As Variant
    Dim x, i As Integer
    x = (GroupList.ListCount - 1)
    ReDim UnitGroupIdX(1 To x), ClientIDx(1 To x), ScheduleIDx(1 To x) As Integer
    For i = 1 To x
         UnitGroupIdX(i) = GroupList.ItemData(i)
         ClientIDx(i) = Me.Parent.clientID
         ScheduleIDx(i) = Me.Parent.SchID
       ' GroupList.Selected(i) = True
    Next i

        If Me.Form.NewRecord = True Then
       ' If Me.Form.Dirty = True Then Me.Form.Dirty = False
       '     Me.clientID = Me.Parent.clientID
       '     Me.ScheduleId = Me.Parent.SchID

    For i = 1 To x

            SQLz = "INSERT INTO BillingRulesClient ( ClientId, ScheduleId, UnitGroupId, PerDiemType, PerDiemStartOn, PerDiemFeeToUse, PerDiemFeeValue, " & _
                   "PerDiemFeeTaxable, PerDiemFeeLabel, PerDiemShowFeeColumn, PerDiemRefundFee, BillType, SplitFee, RemitTo, ShowContract, FeeToUse, FeeValue, " & _
                   "FeeTaxable, FeeTerm, FeeLabel, ShowFeeColumn, StartRentOn, RefundFee, DayOfMonth, CustomStartDate, Script ) " & _
                   "SELECT " & ClientIDx(i) & "," & ScheduleIDx(i) & ", " & UnitGroupIdX(i) & ", " & Me.PerDiemType & ", " & Me.PerDiemStartOn & ", " & Me.PerDiemFeeToUse & ", " & Me.PerDiemFeeValue & ", " & _
                   "" & Nz(Me.PerDiemFeeTaxable, 0) & ", " & Nz(Me.PerDiemFeeLabel, 0) & ", " & Me.PerDiemShowFeeColumn & ", " & Me.PerDiemRefundFee & ", " & Me.BillType & ", " & _
                   "" & Me.SplitFee & ", " & Me.RemitTo & ", " & Me.ShowContract & ", " & Me.FeeToUse & ", " & Me.FeeValue & ", " & _
                   "" & Nz(Me.FeeTaxable, 0) & ", " & Me.FeeTerm & ", " & Me.FeeLabel & ", " & Me.ShowFeeColumn & ", " & Me.StartRentOn & ", " & _
                   "" & Me.RefundFee & ", " & Me.DayOfMonth & ", #" & Nz(Me.CustomStartDate, "1/1/1900") & "#, '" & Replace(Nz(Me.Script, "NO"), "'", """") & "';"

            ' DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
       'Me.GroupList
   Next i
    Me.Undo
       Else
       '   Me.Form.Refresh
        '    SQLz = "UPDATE BillingRulesClient SET  PerDiemType = " & Me.PerDiemType & ", PerDiemStartOn = " & Me.PerDiemStartOn & ", PerDiemFeeToUse = " & Me.PerDiemFeeToUse & ", " & _
                   "PerDiemFeeValue = " & Me.PerDiemFeeValue & ", PerDiemFeeTaxable = " & Me.PerDiemFeeTaxable & ",  PerDiemFeeLabel = " & Me.PerDiemFeeLabel & ", PerDiemShowFeeColumn = " & Me.PerDiemShowFeeColumn & ", " & _
                   "PerDiemRefundFee = " & Me.PerDiemRefundFee & ", BillType = " & Me.BillType & ", SplitFee = " & Me.SplitFee & ", RemitTo = " & Me.RemitTo & ", " & _
                   "ShowContract = " & Me.ShowContract & ", FeeToUse = " & Me.FeeToUse & ", FeeValue = " & Me.FeeValue & ", FeeTaxable = " & Me.FeeTaxable & ", " & _
                   "FeeTerm = " & Me.FeeTerm & ", FeeLabel = " & Me.FeeLabel & ", ShowFeeColumn = " & Me.ShowFeeColumn & ", StartRentOn = " & Me.StartRentOn & ", " & _
                   "RefundFee = " & Me.RefundFee & ",  DayOfMonth = " & Me.DayOfMonth & ",  CustomStartDate = " & Nz(Me.CustomStartDate, "1/1/1900") & ",  Script = " & Nz(Me.Script, "No") & " " & _
                   "WHERE  ClientId=" & Me.clientID & " AND ScheduleId= " & Me.ScheduleId & " AND  UnitGroupId =" & UnitGroupIdX(i) & ""

            ' DoCmd.SetWarnings False
            '       DoCmd.RunSQL (SQLz)
         '   DoCmd.SetWarnings True


      End If
   On Error GoTo FixMe
  ' For i = 1 To x
       '  UnitGroupIdX(i) = GroupList.ItemData(i)
        GroupList.Selected(1) = True
        DoCmd.GotoRecord , , acFirst
        Me.PerdiemFee.SetFocus
   ' Next i
FixMe:
'Set RST = Nothing
End Sub
Private Sub GroupList_Click()
  On Error GoTo FixMe
  If Me.Form.NewRecord = True Then Exit Sub
   Dim RST As DAO.Recordset
   Me.Form.FilterOn = False
   Set RST = Me.RecordsetClone
       RST.FindFirst "UnitGroupId = " & Me.GroupList.Column(0) & ""
       Me.Bookmark = RST.Bookmark

       RST.close
   Set RST = Nothing
FixMe:
End Sub

Private Sub ShowContract_Click()
    Call RunScript
End Sub

Private Sub ShowFeeColumn_Click()
    Call RunScript
End Sub

Private Sub SplitFee_Click()
    If Me.SplitFee = 0 And Me.BillType = 4 Then Me.BillType = 2
    If Me.SplitFee = 0 And Me.BillType = 8 Then Me.BillType = 7
    If Me.FeeToUse = 0 Then
        MsgBox "You first must select a Fee to have Split Billing", vbCritical, "no fee, no split"
        Exit Sub
    End If
    If Me.SplitFee <> 0 Then
        If Me.BillType = 2 Or Me.BillType = 6 Then
            Dim SplitIt As String
            SplitIt = MsgBox("You're trying to use 'Split Billing' for a Non-Split bill type" & vbNewLine & "Do you want change the Bill Type to 'Split Billing'?", vbQuestion + vbYesNo + vbCritical, "I'm confused about your Bill Types")
            If SplitIt = vbNo Then
                Me.SplitFee = 0
                Exit Sub
            Else
                If Me.BillType = 2 Then Me.BillType = 4
                If Me.BillType = 7 Then Me.BillType = 8
            End If
        End If
    End If
    Me.ShowFeeColumn = 0
    Call BillType_Click
    Call RunScript
End Sub

Private Sub StartRentOn_Click()
    If Me.StartRentOn = 3 Then
    Call RunScript
    Me.CustomStartDate.Visible = True
    Me.CustomStartDate.SetFocus
    Else
    Me.CustomStartDate.Visible = False
    Call RunScript
    End If
End Sub
Private Sub CustomStartDate_AfterUpdate()
Call RunScript
End Sub

Private Sub BillType_Click()
    If IsNull(Me.FeeToUse) Or Me.FeeToUse = 0 Or Me.FeeToUse = 5 Then Me.FeeToUse = 0
    If IsNull(Me.FeeTerm) Or Me.FeeTerm = 0 Then Me.FeeTerm = 0
    Select Case Me.BillType
    Case 1          ' Regular Invoice
        If IsNull(Me.FeeToUse) Or Me.FeeToUse = 0 Then Me.FeeToUse = 1
        Me.FeeLabel = 1
        Me.RemitTo = 1
        Me.SplitFee = 0

    Case 2          ' Rent Only
        Me.FeeToUse = 0
        Me.FeeTerm = 0
        Me.RemitTo = 1
        Me.SplitFee = 0
    Case 3          ' Fee Only
        If IsNull(Me.FeeToUse) Or Me.FeeToUse = 0 Then Me.FeeToUse = 1
        Me.RemitTo = 2
        Me.SplitFee = 0
    Case 4          ' split
        If IsNull(Me.FeeToUse) Or Me.FeeToUse = 0 Then Me.FeeToUse = 1
        Me.FeeLabel = 1
        Me.RemitTo = 1
        Me.SplitFee = 1
    Case 5          ' Ancillary
        Me.FeeToUse = 5
        Me.RemitTo = 2
        Me.SplitFee = 0
    Case 6          ' LO Invoice
        If IsNull(Me.FeeToUse) Or Me.FeeToUse = 0 Then Me.FeeToUse = 1
        Me.FeeLabel = 1
        Me.RemitTo = 2
        Me.SplitFee = 0
    Case 7          ' Rent Only LO
        Me.FeeToUse = 0
        Me.FeeTerm = 0
        Me.RemitTo = 2
        Me.SplitFee = 0
    Case 8          ' Split LO
        If IsNull(Me.FeeToUse) Or Me.FeeToUse = 0 Then Me.FeeToUse = 1
        'Me.FeeToUse = 1
        Me.FeeLabel = 1
        Me.RemitTo = 2
        Me.SplitFee = 1
    End Select
    Me.DayOfMonth = 25
    Me.StartRentOn = 1
    Me.FeeTerm = 0
    Me.ShowContract = 1
    Me.ShowFeeColumn = 1
    Me.RefundFee = 0
    Me.FeeTaxable = 0
    'If Me.BillType = 4 Then fFadeControls True, Me, "SplitFee" Else fFadeControls False, Me, "SplitFee"
    Call RemitTo_Click
    Call FeeToUse_Click
    '  Call SplitFee_Click
    Call RunScript

End Sub

Private Sub FeeToUse_Click()
' If Me.BillType = 4 And Me.FeeToUse.Value = 0 Then Exit Sub
    If Me.FeeToUse = 0 Then
        If Me.BillType = 1 Then Me.BillType = 2    ' if Invoice Synd
        If Me.BillType = 4 Then Me.BillType = 2    ' if Split Synd
        If Me.BillType = 6 Then Me.BillType = 7    ' if Invoice LO
        If Me.BillType = 8 Then Me.BillType = 7    ' if Split LO
        Me.ShowFeeColumn = 0
        Me.SplitFee = 0
        Me.RefundFee = 0
        Me.FeeLabel = 0
        If IsNull(Me.FeeTerm) Or Me.FeeTerm = 0 Then Me.FeeTerm = 0
        Call EnableDisableOptionse("Off", "RentFee")
        Call RunScript
        Exit Sub
        '  fFadeControls False, Me, "RentFee"
        '  fFadeControls False, Me, "RentFee2"
        '  fFadeControls False, Me, "RentFee1"
    Else
        If Me.BillType = 2 Or Me.BillType = 7 Then    'rent only types
            MsgBox "There cannot be a fee on a 'RENT ONLY' bill" & vbNewLine & "If you want to have a fee, select another bill type", vbInformation, "Get your 'fees' straight"
            Me.FeeToUse = 0
            Call EnableDisableOptionse("Off", "RentFee")
            Call RunScript
            Exit Sub
        End If
        Call EnableDisableOptionse("On", "RentFee")
        '  fFadeControls True, Me, "RentFee"
        '  fFadeControls True, Me, "RentFee2"
        '  fFadeControls True, Me, "RentFee3"
        Me.ShowFeeColumn = 1
        If Me.BillType <> 4 And Me.BillType <> 8 Then Me.SplitFee = 0 Else Me.SplitFee = 1
        Me.RefundFee = 0
        If IsNull(Me.FeeLabel) Or Me.FeeLabel = 0 Then Me.FeeLabel = Me.FeeToUse
        If IsNull(Me.FeeTerm) Or Me.FeeTerm = 0 Then Me.FeeTerm = 0
    End If
    Call RunScript
End Sub

Private Sub PerDiemFeeToUse_Click()
    If IsNull(Me.PerDiemFeeToUse) Or Me.PerDiemFeeToUse.Column(1) = "None" Then
        Me.PerDiemRefundFee = 0
        Me.PerDiemFeeLabel = Null
        Me.PerDiemShowFeeColumn = 0
        Me.PerDiemFeeTaxable = 0
        Call EnableDisableOptionse("Off", "PDFee")
    Else
        Call EnableDisableOptionse("On", "PDFee")
        Me.PerDiemFeeLabel = Me.PerDiemFeeToUse
        Me.PerDiemShowFeeColumn = 1
        Me.PerDiemRefundFee = 0
    End If
    Call RunScript
End Sub

Private Sub RunPdScript()
    msg = ""
    PDScriptX = ""
    msg = msg & "<Font Size=5><B>" & Me.PerDiemType.Column(1) & "</B></Font><ul>"
    msg = msg & "<li>Per Diem starts on the Unit's<b> " & Me.PerDiemStartOn.Column(1) & " </B>date.</li>"
    If Me.PerDiemType = 0 Then msg = msg & "<li>Per Diem is billed until <b> BLCD Date</b></li>"
    If Me.PerDiemType = 1 Then msg = msg & "<li>Per Diem is billed for<b> 1 Month</b></li>"
    If Me.PerDiemType = 2 Then msg = msg & "<li>Per Diem is billed for<b> 2 Months</b></li>"
    If Me.PerDiemType = 3 Then msg = msg & "<li>Per Diem is billed<b> Just Once </b>for the<b> entire </b>per Diem period</li>"
    If Me.PerDiemType = 4 Then msg = msg & "<li>Per Diem is billed<b> After the 30th day</b></b> of " & Me.PerDiemStartOn.Column(1) & "</li>"
    msg = msg & "<li>The Invoice will remit to<b> Fleet Advantage</b>'s address.</li>"
    msg = msg & "<li>The contact number<b> WILL </b>show on the invoice.</li>"
    If IsNull(Me.PerDiemFeeToUse) Or Me.PerDiemFeeToUse = 0 Then
        msg = msg & "<li>There are<B> No </B>fees for this invoice</li></ul>"
        GoTo Done
    End If

    msg = msg & "Fee Type: <B>" & Me.PerDiemFeeToUse.Column(1) & "</B><ul>"
    msg = msg & "<li>Fee Amount: <B>" & Format(Me.PerdiemFee, "$#,#.#0") & "</B></li>"
    msg = msg & "<li>Billed the<B> ENTIRE </B> term of the Per Diem</li>"
    If Me.PerDiemShowFeeColumn = 0 Then
        msg = msg & "<li>The Fee column will be<b> Hidden & Combined</B> as 'Rent'</li>"
    Else
        msg = msg & "<li>The Fee column will be labeled '<B> " & Me.PerDiemFeeLabel.Column(1) & " </b>'</li>"
    End If
    If Me.PerDiemFeeTaxable <> 0 Then msg = msg & "<li>The Fee is<b> Taxable </B> @ Location's tax rate</li>"
Done:
    msg = msg & "</ul></ul>"
    Me.DescPerDiem = msg
    PDScriptX = msg & "|"
 If IsNull(Me.BillType) Then
 Me.BillType = 1
 Call BillType_Click
 End If
End Sub
Private Sub RunScript()
    Call RunPdScript
    Dim sufx As String
    msg = ""
    BillScriptX = ""
    If Me.BillType = 4 Or Me.BillType = 8 Then
        Me.SplitBox.Visible = True
        If Me.BillType = 4 Then msg = msg & "<Font Size=5><B>Rent Only (Bill 1)</B></Font><ul>"
        If Me.BillType = 8 Then msg = msg & "<Font Size=5><B>LO Rent Only (Bill 1)</B></Font><ul>"
    Else
        Me.SplitBox.Visible = False
        msg = msg & "<Font Size=5><B>" & Me.BillType.Column(1) & "</B></Font><ul>"
    End If
    If Me.StartRentOn = 3 Then
    msg = msg & "<li>Rent starts on: <b> " & Me.CustomStartDate & " </B>.</li>"
    Else
    msg = msg & "<li>Rent starts on the Unit's<b> " & Me.StartRentOn.Column(1) & " </B>date.</li>"
    If IsNull(Me.CustomStartDate) Then CustomStartDate = CDate("1/1/1900")
    End If
    sufx = "th "
        If Me.DayOfMonth = 3 Or Me.DayOfMonth = 23 Then sufx = "rd "
    If Me.DayOfMonth = 2 Or Me.DayOfMonth = 22 Then sufx = "nd "
    If Me.DayOfMonth = 1 Or Me.DayOfMonth = 21 Or Me.DayOfMonth = 31 Then sufx = "st "
    Me.DayofMthLabel.caption = sufx & " Day"
    msg = msg & "<li>The Client is billed on the<b> " & Me.DayOfMonth & sufx & "</b>day of the month</li>"
    If Me.RemitTo = 1 Then
        msg = msg & "<li>The Invoice will remit to<b> Assignee (" & Nz(Me.Remittee, "Bank") & ")</b>'s address.</li>"
    Else
        msg = msg & "<li>The Invoice will remit to<b> Fleet Advantage</b>'s address.</li>"
    End If
    If Me.ShowContract = 1 Then
        msg = msg & "<li>The contract number<b> WILL </b>show on the invoice.</li>"
    Else
        msg = msg & "<li>The contract number<b> will NOT </b>show on the invoice.</li>"
    End If
    'BillScriptX = "|" & Me.BillType.Column(1) & "," & Me.StartRentOn.Column(1) & "," & Me.DayOfMonth & "," & Me.RemitTo & "," & _
     Me.ShowContract.Column(1)
    If IsNull(Me.FeeToUse) Or Me.FeeToUse.Column(1) = "None" Or Me.SplitFee = 1 Then
        If Me.SplitFee = 1 Then
            msg = msg & "<li>This invoice will<B> NOT </B>have a fee, bill <b>2</b> will.</li></ul>"
        Else
            msg = msg & "<li>There are<B> No </B>fees for this invoice</li></ul>"
        End If
        GoTo Done
    End If
    msg = msg & "Fee Type: <B>" & Me.FeeToUse.Column(1) & "</B><ul>"
    msg = msg & "<li>Fee Amount: <B>" & Format(Me.InvoiceFee, "$#,#.#0") & "</B></li>"
    If Me.FeeTerm = 0 Then
        msg = msg & "<li>Billed the<B> ENTIRE </B> term of the lease</li>"
    Else
        msg = msg & "<li>Billed the first<B> " & Me.FeeTerm & " </B>months of the lease</li>"
    End If
    If Me.ShowFeeColumn = 0 Then
        msg = msg & "<li>The Fee column will be<b> Hidden & Combined</B> as 'Rent'</li>"
    Else
        msg = msg & "<li>The Fee column will be labeled '<B> " & Me.FeeLabel.Column(1) & " </b>'</li>"
    End If
    If Me.FeeTaxable <> 0 Then msg = msg & "<li>The Fee is<b> Taxable </B> @ Location's tax rate</li>"
Done:
    msg = msg & "</ul></ul>"
    Me.BillDescRule = msg
    BillScriptX = PDScriptX & msg & "|"
    Me.BillingScript.Value = BillScriptX
    Me.SplitFeeDescRule = ""
    If Me.SplitFee = 1 Then
    Call SplitFeeScript
    Else
   ' Call SaveUpdate_Click
    End If
End Sub
Private Sub SplitFeeScript()
    Dim msg As String
    msg = ""
    msg = msg & "<Font Size=5><B> Fee Only (Bill 2)</B></Font><ul>"
    msg = msg & "<li>This 'Fee-Only' Invoice will remit to<b> Fleet Advantage</b></li>"
    If Me.ShowContract <> 1 Then
        msg = msg & "<li>The contact number<b> will NOT </b>show on the invoice.</li>"
    Else
        msg = msg & "<li>The contact number<b> WILL </b>show on the invoice.</li>"
    End If
    msg = msg & "Fee Type: <B>" & Me.FeeToUse.Column(1) & "</B><ul>"
    msg = msg & "<li>Fee Amount: <B>" & Format(Me.InvoiceFee, "$#,#.#0") & "</B></li>"
    If Me.FeeTerm = 0 Then
        msg = msg & "<li>Billed the<B> ENTIRE </B> term of the lease</li>"
    Else
        msg = msg & "<li>Billed the first<B> " & Me.FeeTerm & " </B>months of the lease</li>"
    End If
    If Me.FeeTaxable <> 0 Then msg = msg & "<li>The Fee is<b> Taxable </B> @ Location's tax rate</li>"
    msg = msg & "</ul></ul>"
    Me.SplitFeeDescRule = msg
    BillScriptX = BillScriptX & msg & "|"
    Me.BillingScript = BillScriptX
   Call SaveUpdate_Click
End Sub

Private Sub RemitTo_Click()
'If Me.RemitTo = 1 Then Me.Remittee = DLookup("Assignee_Bank", "SchGrp", "SGrpID=" & Me.[UnitGroupId] & "") Else Me.Remittee = "Fleet"
    Call RunScript
End Sub
