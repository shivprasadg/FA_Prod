Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim CityType   As String

Private Sub Form_Load()
    Me.SelectCountry = "US"
    Me.SelectState = 10
    Me.StateID = 10
   ' Me.CountyID = 326
    Me.CityID = 0
    Call SelectCountry_Click
    Call SelectState_Click
End Sub

Public Sub SelectCountry_Click()
    Me.CountryID = Me.SelectCountry.Column(0)
    SQLz = "SELECT id, StateName, PostalAbbreviation AS State, SalesTaxOrUseTaxNumber, ISOCountryCode2 " & _
           "FROM TaxStates WHERE ISOCountryCode2='USA' Order by PostalAbbreviation"
    Me.SelectState.RowSource = SQLz

End Sub


Public Sub SelectState_Click()
    Me.StateID = Me.SelectState
    Me.CityID = 0
    Me.CountyID = 0
    Me.SelectCity = 0
    'Me.SelectCounty = 0
    ' Sqlz = "SELECT id, CountyName, TaxStatesID FROM TaxCounties Where TaxStatesID=" & Me.StateID & " ORDER BY CountyName;"
    ' Me.SelectCounty.RowSource = Sqlz
    SQLz = "SELECT ID, CityName, TaxStateID " & _
           "FROM TaxCities WHERE TaxStateID =" & Me.StateID & " ORDER BY CityName;"
    Me.SelectCity.RowSource = SQLz
    CityType = "City"
    Me.StateLabel2.caption = CityType & Me.SelectCity.Column(1) & " Detail"
    Me.StateLabel3.caption = CityType & Me.SelectCity.Column(1) & " Taxes"
    Me.StateLabel4.caption = CityType & Me.SelectCity.Column(1) & " Exceptions"
End Sub
Public Sub SelectCity_Click()
    Me.CityID = Me.SelectCity
    'Me.SelectCounty = Me.SelectCity.Column(2)
    'Me.CountyID = Me.SelectCity.Column(2)
    CityType = "City of "
    Me.StateLabel2.caption = CityType & Me.SelectCity.Column(1) & " Detail"
    Me.StateLabel3.caption = CityType & Me.SelectCity.Column(1) & " Taxes"
    Me.StateLabel4.caption = CityType & Me.SelectCity.Column(1) & " Exceptions"
End Sub
Private Sub AddTaxRate_Click()
    If Nz(Me.SelectCity, 0) = 0 Then
        MsgBox "You must Select a City", vbCritical, AtlaasVersion
        Exit Sub
    End If
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TaxBaseTemp")
    DoCmd.SetWarnings True
    DoCmd.close acForm, "Tax_Add"
    DoCmd.OpenForm "Tax_Add", acNormal, , , , , "City|" & Nz(Me.TaxTypeId, 0) & "|" & Me.CityID & "|0|Add"
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub
