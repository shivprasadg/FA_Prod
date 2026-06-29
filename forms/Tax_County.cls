Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Dim CountyType As String
Dim CurrLevelTxRate As Double



Private Sub AddTaxRate_Click()
    If Me.SelectCounty = 0 Or IsNull(Me.SelectCounty) Then
        MsgBox "You must Select a County", vbCritical, AtlaasVersion
        Exit Sub
    End If
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TaxBaseTemp")
    DoCmd.SetWarnings True
    DoCmd.close acForm, "Tax_Add"
    DoCmd.OpenForm "Tax_Add", acNormal, , , , , "County|" & Nz(Me.Tax_CountyBaseSub!TaxTypeID, 0) & "|" & Me.CountyID & "|0|Add"
End Sub
Private Sub CloseMe_Click()
DoCmd.close acForm, Me.Form.Name
End Sub
Private Sub Form_Load()
Me.SelectCountry = "US"
Me.SelectState = 10
Me.StateID = 0
Me.CountyID = 0
Call SelectCountry_Click
Call SelectState_Click
End Sub

Private Sub SelectCountry_Click()
Me.CountryID = Me.SelectCountry.Column(0)
If Me.SelectCountry = "CA" Then Me.StateLabel1.caption = "Province" Else Me.StateLabel1.caption = "State"
SQLz = "SELECT id, StateName, PostalAbbreviation AS State, SalesTaxOrUseTaxNumber, ISOCountryCode2 " & _
       "FROM TaxStates WHERE ISOCountryCode2='USA' Order by PostalAbbreviation"
Me.SelectState.RowSource = SQLz
Me.CountyID = 0
Me.SelectCounty = 0
Me.SelectState = 0
Me.StateID = 0
CountyType = " County "
If Me.StateID = 2 Or Me.StateID = 9 Then CountyType = " "
Me.StateLabel2.caption = Me.SelectCounty.Column(1) & CountyType & "Detail"
Me.StateLabel3.caption = Me.SelectCounty.Column(1) & CountyType & "Taxes"

End Sub
Private Sub SearchCounty_Change()
'SQlz = "Select * From TaxCounties WHERE TaxStatesID=" & Me.StateID & " AND CountyName Like '*" & Me.SearchCounty.Text & "*' Order By CountyName"
'Me.Tax_CountySub.Form.RecordSource = SQlz
End Sub

Public Sub SelectCounty_Click()
Me.CountyID = Me.SelectCounty
CountyType = " County "
If Me.StateID = 2 Or Me.StateID = 19 Then CountyType = " "
Me.StateLabel2.caption = Me.SelectCounty.Column(1) & CountyType & "Detail"
Me.StateLabel3.caption = Me.SelectCounty.Column(1) & CountyType & "Taxes"
Me.StateLabel4.caption = Me.SelectCounty.Column(1) & CountyType & "Exceptions"
End Sub

Public Sub SelectState_Click()
Me.StateID = Me.SelectState
SQLz = "SELECT id, CountyName, TaxStatesID FROM TaxCounties Where TaxStatesID=" & Me.StateID & " ORDER BY CountyName;"
Me.SelectCounty.RowSource = SQLz
Me.CountyID = 0
Me.SelectCounty = 0
Dim CountyType As String
If Me.StateID <> 2 Or Me.StateID <> 19 And Me.CountryID = "US" Then CountyType = " " Else CountyType = " County "
Me.StateLabel2.caption = Me.SelectCounty.Column(1) & CountyType & "Detail"
Me.StateLabel3.caption = Me.SelectCounty.Column(1) & CountyType & "Taxes"
Me.StateLabel4.caption = Me.SelectCounty.Column(1) & CountyType & "Exceptions"
End Sub
