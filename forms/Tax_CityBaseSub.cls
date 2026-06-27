Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim CheckUpdate As String
Dim CurrentRate As Double
Dim Response   As String

Private Sub Form_Current()
    On Error GoTo FixMe
    Me.txtCurrent = Me.ID
    If Not IsNull(Me.TaxTypeId) Then Me.Parent.TaxTypeId = Me.TaxTypeId
FixMe:
End Sub
Private Sub EditMe_Click()
    SQLz = ""
    DoCmd.close acForm, "Tax_Add"
    SQLz = SQLz & "INSERT INTO TaxBaseTemp("
    SQLz = SQLz & " TaxLevelLookupID"
    SQLz = SQLz & ",TaxTypeID"
    SQLz = SQLz & ",BaseTaxRate"
    SQLz = SQLz & ",StartsOn"
    SQLz = SQLz & ",EndsOn"
    SQLz = SQLz & ",TaxLevel"
    SQLz = SQLz & ",EditedTaxID"
    SQLz = SQLz & ",TaxScript"
    SQLz = SQLz & ") SELECT "
    SQLz = SQLz & " TaxCitiesID"
    SQLz = SQLz & ",TaxTypeID"
    SQLz = SQLz & ",BaseTaxRate"
    SQLz = SQLz & ",StartsOn"
    SQLz = SQLz & ",EndsOn"
    SQLz = SQLz & ",'City' As TaxLevel"
    SQLz = SQLz & ",ID"
    SQLz = SQLz & ",TaxScript"
    SQLz = SQLz & " FROM TaxCitiesBase"
    SQLz = SQLz & " WHERE ID =" & Me.ID & ";"

    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TaxBaseTemp")
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    DoCmd.OpenForm "Tax_Add", acNormal, , , , , "City|" & Me.Parent.TaxTypeId & "|" & Me.Parent.CityID & "|" & CurrLevelTxRate & "|Edit"
End Sub
Public Sub TaxNumber_GotFocus()
    CurrentRate = Me.TaxNumber
End Sub
