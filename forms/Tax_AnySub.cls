Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Private Sub Form_Current()
On Error GoTo FixMe
    Me.txtCurrent = Me.ID
FixMe:
End Sub

Private Sub Form_Load()
Dim SourceX As String
Dim TypeIDx As String
TypeIDx = "TaxStatesID"
SourceX = "vw_Tax_StateBaseQ"
If Me.Parent.TaxLevel = "City" Then TypeIDx = "TaxCityID"
If Me.Parent.TaxLevel = "City" Then SourceX = "vw_Tax_CityBaseQ"
If Me.Parent.TaxLevel = "County" Then TypeIDx = "TaxCountyID"
If Me.Parent.TaxLevel = "County" Then SourceX = "vw_Tax_CountyBaseQ"
SQLz = "Select * From " & SourceX & " WHERE " & TypeIDx & "=" & Me.Parent.TaxLevelLookupID & " Order By StartsOn Desc;"
Me.Form.RecordSource = SQLz
End Sub

Private Sub TaxDescription_GotFocus()
'Me.TaxDescription.ControlTipText = Me.TaxDescription
End Sub
