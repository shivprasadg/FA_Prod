Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private ranCurrent As Boolean

Private Sub EditState_Click()
    If Me.Parent.AddressText Like "*ZZ*" Then
        MsgBox "In Order to Modify or Add a State tax" & vbNewLine & "You must select a valid State for this location.", vbInformation, AtlaasVersion
        Me.Parent.LState.SetFocus
        Exit Sub
    End If
    DoCmd.close acForm, "Tax_State"
    DoCmd.OpenForm "Tax_State"
    Forms!Tax_State.SelectState = Me.StateID
    Forms!Tax_State.StateID = Me.StateID
  '  Forms!Tax_State!Tax_StateSub.Form.GotoRecord (Me.StateID)
End Sub
Private Sub EditCounty_Click()
    If Me.Parent.AddressText Like "*County Unknown*" Then
        MsgBox "In Order to Modify or Add a County tax" & vbNewLine & "You must select a valid County for this location.", vbInformation, AtlaasVersion
        Me.Parent.LCounty.SetFocus
        Exit Sub
    End If
    DoCmd.close acForm, "Tax_County"
    DoCmd.OpenForm "Tax_County"
    Forms!Tax_County.StateID = Me.StateID
    Forms!Tax_County.SelectState = Me.StateID
    Forms!Tax_County.SelectState_Click
    Forms!Tax_County.CountyID = Me.CountyID
    Forms!Tax_County.SelectCounty = Me.CountyID
    Forms!Tax_County.SelectCounty_Click
'    Forms!Tax_County!Tax_CountySub.Form.GotoRecord (Me.CountyID)
End Sub
Private Sub EditCity_Click()
If Me.Parent.AddressText Like "*Unknown,*" Then
        MsgBox "In Order to Modify or Add a City tax" & vbNewLine & "You must select a valid City for this location.", vbInformation, AtlaasVersion
        'Me.Parent.LCity.SetFocus
        Exit Sub
    End If
    DoCmd.close acForm, "Tax_City"
    DoCmd.OpenForm "Tax_City"
    Forms!Tax_City.SelectCountry_Click
    Forms!Tax_City.SelectState = Me.StateID
    Forms!Tax_City.SelectState_Click
   ' Forms!Tax_City.SelectCounty = Me.CountyID
    'Forms!Tax_City.SelectCounty_Click
    Forms!Tax_City.SelectCity = Me.CityID
    Forms!Tax_City.SelectCity_Click
End Sub
Private Sub Form_Current()
 If ranCurrent = True Then Exit Sub     ' prevents double execution
    ranCurrent = True
 Dim Ccrl As Control
    For Each Ccrl In Me.Controls
 '   If Me(ctrl.Name).Type = TextBox Then
   '     If Me.SalesTax = 0 And Me(cCrl.Name).Tag = "Sales" Then Me(cCrl.Name).Height = 0 Else Me(cCrl.Name).Height = 288
        If Me.AdValorem = 0 And Me(Ccrl.Name).Tag = "AdValorem" Then Me(Ccrl.Name).Height = 0
        If Me.AdValorem > 0 And Me(Ccrl.Name).Tag = "AdValorem" Then Me(Ccrl.Name).Height = 278
        If Me.FHUTTax = 0 And Me(Ccrl.Name).Tag = "FHUT" Then Me(Ccrl.Name).Height = 0
        If Me.FHUTTax > 0 And Me(Ccrl.Name).Tag = "FHUT" Then Me(Ccrl.Name).Height = 278
        If Me.HUTTax = 0 And Me(Ccrl.Name).Tag = "HUT" Then Me(Ccrl.Name).Height = 0
        If Me.HUTTax > 0 And Me(Ccrl.Name).Tag = "HUT" Then Me(Ccrl.Name).Height = 278
  '  End If
    Next Ccrl
    'If Nz(Me.Parent.LTaxRate, 0) = 0 Then Me.Parent.LTaxRate = Val((Me.[StateSales] + Me.[CountySales] + Me.[CitySales]))
 '   If Nz(Me.Parent.LTaxRate, 0) <> ([StateSales] + [CountySales] + [CitySales]) Then Me.Parent.LTaxRate.BackColor = RGB(255, 204, 0) Else Me.Parent.LTaxRate.BackColor = RGB(215, 228, 189)
End Sub

Private Sub SaveUpdate_Click()
If Me.Dirty Then Me.Dirty = False
Me.Form.Requery
End Sub
