Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Private Sub Form_Load()
Me.SelectCountry = "US"
Me.StateID = "0"
Me.CountyID = "0"
Call SelectCountry_Click
End Sub
Private Sub AddException_Click()
'If Me.SelectState.Column(1) = "Unknown" Then
'Call Dialog.RichBox("<font size=3>'Unknown' State</font><br />You<b> must</b> first select a valid state<br />Before you can add an exception", vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
'Exit Sub
'End If
'    Dim CurrLevelTxRate As Double
'    If Me.Tax_StateBaseSub.Form.RecordsetClone.RecordCount > 0 Then
'        CurrLevelTxRate = Nz(Me.Tax_StateBaseSub!TaxNumber, 0)
'    Else
'        CurrLevelTxRate = 0
'        MsgBox "This State: " & Me.SelectState.Column(1) & " dosen't have a base tax." & vbNewLine & "You cannot have an exception without a 'base' tax", vbCritical, AtlaasVersion
'        Exit Sub
'    End If
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL ("Delete From TaxExceptionsTemp")
'    DoCmd.SetWarnings True
'    DoCmd.close acForm, "Tax_Add_Exception"
'    DoCmd.OpenForm "Tax_Add_Exception", acNormal, , , , , "State|" & Nz(Me.Tax_StateBaseSub!TaxTypeId, 1) & "|" & Me.StateID & "|" & CurrLevelTxRate & "|Add"
End Sub

Private Sub AddTaxRate_Click()
    If Me.SelectState.Column(1) = "Unknown" Then
        Call Dialog.RichBox("<font size=3>'Unknown' State</font><br />You<b> must</b> first select a valid state<br />Before you can add a Tax", vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
        Exit Sub
    End If
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TaxBaseTemp")
    DoCmd.SetWarnings True
    DoCmd.close acForm, "Tax_Add"
    DoCmd.OpenForm "Tax_Add", acNormal, , , , , "State|" & Nz(Me.Tax_StateBaseSub!TaxTypeID, 1) & "|" & Me.StateID & "|0|Add"
End Sub

Private Sub SelectCountry_Click()
Me.CountryID = Me.SelectCountry.Column(0)
If Me.SelectCountry = "CA" Then
Me.Tax_StateSub!StateNameLabel.caption = "Province"
Me.StateLabel2.caption = "Province Detail"
Me.StateLabel3.caption = "Provincial Taxes"
Me.StateLabel4.caption = "Provincial Exceptions"
Me.SelectState = 174
Me.StateID = 174
Me.Tax_StateSub!TaxNumberLabel.caption = "Province"
Else
Me.Tax_StateSub!StateNameLabel.caption = "State Detail"
Me.StateLabel2.caption = "State Detail"
Me.StateLabel3.caption = "State Taxes"
Me.StateLabel4.caption = "State Exceptions"
Me.Tax_StateSub!TaxNumberLabel.caption = "State"
Me.SelectState = 0
Me.StateID = 0
End If
SQLz = "SELECT id, StateName, PostalAbbreviation AS State, SalesTaxOrUseTaxNumber, ISOCountryCode2 " & _
           "FROM TaxStates WHERE ISOCountryCode2='USA' Order by PostalAbbreviation"
Me.SelectState.RowSource = SQLz

'SQlz = "Select * From TaxStates Where ISOCountryCode2 ='" & Me.SelectCountry & "' Order by PostalAbbreviation"
'Me.Tax_StateSub.Form.RecordSource = SQlz

End Sub

Private Sub SelectState_Click()
Me.StateID = Me.SelectState
Me.StateLabel2.caption = Me.SelectState.Column(1) & " Detail"
Me.StateLabel3.caption = Me.SelectState.Column(1) & " Taxes"
Me.StateLabel4.caption = Me.SelectState.Column(1) & " Exceptions"
End Sub
Private Sub CloseMe_Click()
DoCmd.close acForm, Me.Form.Name
End Sub
