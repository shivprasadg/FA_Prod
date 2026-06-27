Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Private Sub Form_Current()
On Error GoTo FixMe
    Me.txtCurrent = Me.ID
    Me.Parent.CityID = Me.ID
    Me.Parent.CountyID = Me.TaxCountyId
FixMe:
End Sub
