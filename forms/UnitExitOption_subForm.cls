Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Compare Database
'
'Private Sub ClearIt_Click()
''Me.ProfileTypeID = 0
''Me.RentActive = 0
''Me.ServiceFee = 0
''Me.EffectiveDate = Null
''Me.EOLDate = Null
''Me.ExtendTermOption = Null
''Call ExtendOrTerm
'End Sub
'
'Private Sub ExchangedByUnitId_DblClick(Cancel As Integer)
''DoCmd.close acForm, "UnitSnapShot"
''DoCmd.OpenForm "UnitSnapShot", acNormal, , "UnitID = " & Me.ExchangedByUnitId & ""
'End Sub
'
'
'
'Private Sub EffectiveDate_AfterUpdate()
''Call LeaseExtTerm_Click
'End Sub
'
'Private Sub ExchangeITType_Click()
'
' '   Me.LabelYes.BackColor = RGB(255, 255, 255)
' '   Me.LabelYes.ForeColor = RGB(31, 73, 125)
' '   Me.ExchangeITMatch.Enabled = False
'    ''Me.LabelNo.BackColor = RGB(255, 255, 255)
'    ''Me.LabelNo.ForeColor = RGB(31, 73, 125)
''CheckIT:
'    'If Me.ExchangeITType = -1 Then
'  '      Me.ExchangeITMatch.Enabled = True
'  '      Me.LabelYes.ForeColor = RGB(255, 255, 255)
'  '      Me.LabelYes.BackColor = RGB(31, 73, 125)
'    'Else
'      '  If Me.ExchangeITType = 0 And Me.ExchangeITMatch & "test" <> "test" Then
'      '      Select Case MsgBox("This Unit has already been Matched, " _
'      '                       & vbCrLf & "You're about to remove that match. " _
'      '                       & vbCrLf & "Are you Sure?" _
'      '           , vbYesNo Or vbExclamation Or vbDefaultButton1, "Remove ExchangeIT Match")
'      '      Case vbYes
'      '          Me.ExchangeITMatch = ""
'      '      Case vbNo
'      '          Me.ExchangeITType = -1
'      '          GoTo CheckIT
'      '          Exit Sub
'      '      End Select
'      '  End If
'
'    'End If
'    ''If Me.ExchangeITType = 0 Then
'    ''Me.ExchangeITSch.Enabled = False
'    ''Me.LabelNo.ForeColor = RGB(255, 255, 255)
'    ''Me.LabelNo.BackColor = RGB(31, 73, 125)
'    ''Me.ExchangeITSch = ""
'    ''End If
'End Sub
'
'Private Sub ExtendTermOption_Click()
'Call LeaseExtTerm_Click
'End Sub
'
'Private Sub Form_Load()
'Call ExtendOrTerm
'Call ExchangeITType_Click
'End Sub
'Private Sub Form_Current()
'Call ExtendOrTerm
'Call ExchangeITType_Click
'End Sub
'
'Private Sub LeaseExtTerm_Click()
'If IsDate(EffectiveDate) Then
'Me.EOLDate = DateAdd("m", Me.ExtendTermOption, EffectiveDate)
'End If
'End Sub
'
'Private Sub ProfileTypeID_Click()
'Call ExtendOrTerm
'Me.RentActive = 0
'Me.ServiceFee = 0
'If ProfileTypeID = 1 Then Me.RentActive = Nz(Me.Parent.FirstExtRent, 0)
'If ProfileTypeID = 2 And Nz(Me.Parent.MTMRent, 0) = 0 Then Me.RentActive = Nz(Me.Parent.UNITRENT, 0)
'If ProfileTypeID = 2 And Nz(Me.Parent.MTMRent, 0) <> 0 Then Me.RentActive = Nz(Me.Parent.MTMRent, 0)
'If ProfileTypeID = 3 Then Me.RentActive = Nz(Me.Parent.UNITRENT, 0)
'End Sub
'
'Private Sub ExtendOrTerm()
'Exit Sub
'' Turns on of off Options based on Lease Option selected
' '   Dim cCrl As control
'    Dim TagOn, TagOff As String
'    TagOn = ""
'    TagOff = ""
'    If Me.ProfileTypeID >= 1 Then
'        If Me.ProfileTypeID >= 1 And Me.ProfileTypeID <= 3 Then
'            ' If the lease is set to Extend turn on Extend Options
'            TagOn = "Extend"
'            TagOff = "TurnIn"
'        End If
'        If Me.ProfileTypeID >= 4 And Nz(Me.ProfileTypeID, 0) <> 0 Then
'            ' If the lease is set to End turn on TurnIn Options
'            TagOn = "TurnIn"
'            TagOff = "Extend"
'        End If
'            If Me.ProfileTypeID = 5 Then
'                Me.EarlyTermFee.Visible = True
'                Me.EarlyTermFeeLabel.Visible = True
'            Else
'                Me.EarlyTermFee.Visible = False
'                Me.EarlyTermFeeLabel.Visible = False
'                Me.EarlyTermFee = 0
'            End If
'    End If
'    For Each Ccrl In Me.Controls
'        ' dis-enbale the contols
'        If Me(Ccrl.Name).Tag = "Notice" Then Me(Ccrl.Name).Enabled = False
'        If Me(Ccrl.Name).Tag = "Extend" Then Me(Ccrl.Name).Enabled = False
'        If Me(Ccrl.Name).Tag = "TurnIn" Then Me(Ccrl.Name).Enabled = False
'        ' enable dis-enable the affected Controls
'        If TagOn <> "" Then
'            If Me(Ccrl.Name).Tag = TagOn Then Me(Ccrl.Name).Enabled = True
'            If Me(Ccrl.Name).Tag = TagOff Then Me(Ccrl.Name).Enabled = False
'            If Me(Ccrl.Name).Tag = "Notice" Then Me(Ccrl.Name).Enabled = True
'        End If
'        ' Set Label colors to defaults
'        If Me(Ccrl.Name).Tag = "OptionLabels" Then Me(Ccrl.Name).BackColor = RGB(255, 255, 255)
'        If Me(Ccrl.Name).Tag = "OptionLabels" Then Me(Ccrl.Name).ForeColor = RGB(31, 73, 125)
'    Next Ccrl
'    ' Highlight the selected labels accoringly
'    If Me.ProfileTypeID = 1 Then Me.Label1.BackColor = RGB(34, 177, 76)
'    If Me.ProfileTypeID = 1 Then Me.Label1.ForeColor = RGB(255, 255, 255)
'    If Me.ProfileTypeID = 2 Then Me.Label2.BackColor = RGB(34, 177, 76)
'    If Me.ProfileTypeID = 2 Then Me.Label2.ForeColor = RGB(255, 255, 255)
'    If Me.ProfileTypeID = 3 Then Me.Label3.BackColor = RGB(34, 177, 76)
'    If Me.ProfileTypeID = 3 Then Me.Label3.ForeColor = RGB(255, 255, 255)
'    If Me.ProfileTypeID = 4 Then Me.Label4.BackColor = RGB(255, 194, 15)
'    If Me.ProfileTypeID = 5 Then Me.Label5.BackColor = RGB(255, 194, 15)
'End Sub
'
'Private Sub SubmitAssigneeNotice_Click()
'MsgBox "This is were the Notice Process would happen." & vbNewLine & "I will set the Sent date to today", , "Notice Sent"
'Me.NoticeSentAssignee = Date
'End Sub
