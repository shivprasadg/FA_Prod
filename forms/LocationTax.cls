Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private ranCurrent As Boolean
Private mLastID As Variant   ' module-level

Private Sub EditState_Click()
    If Nz(Me.Parent.LState, "XX") = "XX" Then
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
    If Nz(Me.Parent.LCounty, "Unknown") = "Unknown" Then
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
If Nz(Me.Parent.LCity, "Unknown") = "Unknown" Then
        MsgBox "In Order to Modify or Add a City tax" & vbNewLine & "You must select a valid City for this location.", vbInformation, AtlaasVersion
        Me.Parent.LCity.SetFocus
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
    If Me.NewRecord Then Exit Sub            ' Skip if no record yet
    If mLastID = Me.LocationID Then Exit Sub ' If we're on the same record as last time, bail out
    mLastID = Me.LocationID

 Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me.AdValorem = 0 And Me(Ccrl.Name).Tag = "AdValorem" Then Me(Ccrl.Name).Height = 0
        If Me.AdValorem > 0 And Me(Ccrl.Name).Tag = "AdValorem" Then Me(Ccrl.Name).Height = 278
        If Me.FHUTTax = 0 And Me(Ccrl.Name).Tag = "FHUT" Then Me(Ccrl.Name).Height = 0
        If Me.FHUTTax > 0 And Me(Ccrl.Name).Tag = "FHUT" Then Me(Ccrl.Name).Height = 278
        If Me.HUTTax = 0 And Me(Ccrl.Name).Tag = "HUT" Then Me(Ccrl.Name).Height = 0
        If Me.HUTTax > 0 And Me(Ccrl.Name).Tag = "HUT" Then Me(Ccrl.Name).Height = 278
    Next Ccrl
    Dim sTax, oTax As Variant
    sTax = (Nz([StateSales], 0#) + Nz([CountySales], 0#) + Nz([CitySales], 0#))
    oTax = Nz(Me.Parent.LTaxRate, 0#) - sTax
    Me.OtherTax = oTax

    If oTax <> 0 Then
       Me.OtherTax.Visible = True
       Me.OtherTaxLabel.Visible = True
    Else
       Me.OtherTax.Visible = False
       Me.OtherTaxLabel.Visible = False
    End If
 End Sub

Private Sub SaveUpdate_Click()
If Me.Dirty Then Me.Dirty = False
Me.Form.Requery
End Sub
