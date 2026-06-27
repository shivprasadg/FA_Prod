Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim SourceX    As String
Dim TargetIDx  As Integer
Dim TaxLevelIDx As String
Dim LevelX     As String

Private Sub Form_Current()
On Error GoTo FixMe
    Me.txtCurrent = Me.ID
FixMe:
End Sub

Private Sub ExceptionDescription_GotFocus()
Me.ExceptionDescription.ControlTipText = Left(RemoveRichText(Nz(Me.TaxScript, "Description")), 252) & "..."
End Sub

Private Sub Form_Load()
SetFormLevel
End Sub
Public Sub SetFormLevel()
    SourceX = "TaxStatesExceptions"
    TaxLevelIDx = "TaxStatesID"
    LevelX = "State"
    CurrentRate = 0
    On Error GoTo FixMe
    If Me.Parent.Form.Name Like "*Cit*" Then
        SourceX = "TaxCitiesExceptionsQ"
        TaxLevelIDx = "TaxCitiesID"
        LevelX = "City"
    '    TargetIDx = Me.Parent.CityID
    '    CurrentRate = Nz(Me.Parent!Tax_CityBaseSub.Form!TaxNumber, 0)
    End If
    If Me.Parent.Form.Name Like "*State*" Then
        SourceX = "TaxStatesExceptionsQ"
        TaxLevelIDx = "TaxStatesID"
        LevelX = "State"
        'TargetIDx = Me.Parent.StateID
     '   CurrentRate = Nz(Me.Parent!Tax_StateBaseSub.Form!TaxNumber, 0)
    End If
    If Me.Parent.Form.Name Like "*Count*" Then
        SourceX = "TaxCountiesExceptionsQ"
        TaxLevelIDx = "TaxCountiesID"
        LevelX = "County"
        'TargetIDx = Me.Parent.CountyID
        'CurrentRate = Nz(Me.Parent!Tax_CountyBaseSub.Form!TaxNumber, 0)
    End If
    SQLz = "Select * From " & SourceX & "Order By ID Desc"
Me.Form.RecordSource = SQLz
FixMe:
End Sub
