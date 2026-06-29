Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim SourceX    As String
Dim SourceX2    As String
Dim TargetIDx  As Integer
Dim TaxLevelIDx As String
Dim TaxLevelIDx2 As String
Dim LevelX     As String
Dim ModeX      As String
Dim CheckUpdate As String
Dim CurrentRate As Double
Public Sub SetFormLevel()
    'On Error GoTo FixMe
    'CurrentRate = Nz(Me.TaxNumber, 0)
    If Me.Parent.Form.Name = "Tax_City" Then
        SourceX = "TaxCitiesBase"
        SourceX2 = "TaxCities"
        TaxLevelIDx = "TaxCityID"
        LevelX = "City"
    '    TargetIDx = Me.Parent.CityID
    End If
    If Me.Parent.Form.Name = "Tax_State" Then
        SourceX = "TaxStatesBase"
        SourceX2 = "TaxStates"
        TaxLevelIDx = "TaxStatesID"
        LevelX = "State"
      '  TargetIDx = Me.Parent.StateID
    End If
    If Me.Parent.Form.Name = "Tax_County" Then
        SourceX = "TaxCountiesBase"
        SourceX2 = "TaxCounties"
        TaxLevelIDx = "TaxCountyID"
        LevelX = "County"
    '    TargetIDx = Me.Parent.CountyID
    End If
FixMe:
End Sub

Private Sub AddExemption_Click()
'If Me.SelectState.Column(1) = "Unknown" Then
'Call Dialog.RichBox("<font size=3>'Unknown' State</font><br />You<b> must</b> first select a valid state<br />Before you can add an exception", vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
'Exit Sub
'End If
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TaxExceptionsTemp")
    DoCmd.SetWarnings True
    DoCmd.close acForm, "Tax_Add_Exception"
    DoCmd.OpenForm "Tax_Add_Exception", acNormal, , , , , LevelX & "|" & Me.TaxTypeID & "|" & Me.LinkID & "|" & Nz(Me.TaxNumber, 0) & "|Add"
'    DoCmd.OpenForm "Tax_Add_Exception", acNormal, , , , , "State|" & Nz(Me.Tax_StateBaseSub!TaxTypeId, 1) & "|" & Me.StateID & "|" & CurrLevelTxRate & "|Add"

End Sub

Private Sub EditMe_Click()
If Me.StateNameLabel = "Unknown" Then
Call Dialog.RichBox("<font size=3>'Unknown' " & SplitString(Me.Parent.Form.Name, "Tax_", 1) & " </font><br />You<b> must</b> first select a valid " & SplitString(Me.Parent.Form.Name, "Tax_", 1) & "<br />Before you can edit", vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
Exit Sub
End If
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
    SQLz = SQLz & ",TaxNote"
    SQLz = SQLz & ",TaxScript"
    SQLz = SQLz & ") SELECT "
    If LevelX = "City" Then SQLz = SQLz & " TaxCityID"
    If LevelX = "State" Then SQLz = SQLz & " TaxStatesID"
    If LevelX = "County" Then SQLz = SQLz & " TaxCountyID"
    SQLz = SQLz & ",TaxTypeID"
    SQLz = SQLz & ",BaseTaxRate"
    SQLz = SQLz & ",StartsOn"
    SQLz = SQLz & ",EndsOn"
    SQLz = SQLz & ",'" & LevelX & "' As TaxLevel"
    SQLz = SQLz & ",ID"
    SQLz = SQLz & ",TaxNote"
    SQLz = SQLz & ",TaxScript"
    SQLz = SQLz & " FROM " & SourceX & ""
    SQLz = SQLz & " WHERE ID =" & Me.ID & ";"

    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TaxBaseTemp")
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    DoCmd.OpenForm "Tax_Add", acNormal, , , , , LevelX & "|" & Me.Parent.TaxTypeID & "|" & Me.LinkID & "|" & Me.TaxNumber & "|Edit"
End Sub

Private Sub Form_Current()
    On Error GoTo FixMe
    Me.txtCurrent = Me.ID
    Me.Parent.TaxTypeID = Me.TaxTypeID
 '   Call SetFormLevel
FixMe:
End Sub

Private Sub Form_Load()
    Call SetFormLevel
    Dim TaxLevelField
    SQLz = ""
    SQLz = SQLz & "SELECT "
    If LevelX = "City" Then SQLz = SQLz & "TaxCities.CityName As BaseName"
    If LevelX = "State" Then SQLz = SQLz & "TaxStates.StateName As BaseName"
    If LevelX = "County" Then SQLz = SQLz & "TaxCounties.CountyName As BaseName"
    If LevelX = "City" Then TaxLevelField = "TaxCities"
    If LevelX = "State" Then TaxLevelField = "TaxStates"
    If LevelX = "County" Then TaxLevelField = "TaxCounties"
    SQLz = SQLz & ", " & SourceX & ".*"
    SQLz = SQLz & ", " & SourceX & "." & TaxLevelIDx & " As LinkID"
    SQLz = SQLz & ",TaxType.TaxAbbreviation"
    SQLz = SQLz & ",TaxType.TaxDescription"
    SQLz = SQLz & ",IIf([StartsOn] > date(),0,IIf([StartsOn] <= date() And nz([EndsOn],cdate('1/1/2050')) >=date(),-1,0)) AS IsActive"
    SQLz = SQLz & ",TaxType.TaxFullName"
    SQLz = SQLz & ",TaxStatus('" & LevelX & "',TaxTypeID," & TaxLevelField & ".ID) As TaxStatus"
    SQLz = SQLz & " FROM (" & SourceX & " INNER JOIN " & SourceX2 & " ON " & SourceX & "." & TaxLevelIDx & "= " & SourceX2 & ".Id) "
    SQLz = SQLz & " LEFT JOIN TaxType ON"
    SQLz = SQLz & " " & SourceX & ".TaxTypeId = TaxType.Id"
    SQLz = SQLz & " ORDER BY "
    SQLz = SQLz & " IIf(Not IsNull([StartsOn]) And IsNull([EndsOn]),-1,0) DESC , TaxType.ID, StartsOn DESC "
    ' SQlz = SQlz & "IIf(Not IsNull([StartsOn]) And IsNull([EndsOn]),-1,0) DESC "
    SQLz = SQLz & ", TaxType.TaxFullName DESC;"
Me.Form.RecordSource = SQLz
End Sub


Public Sub TaxNumber_GotFocus()
    CurrentRate = Me.TaxNumber
End Sub
