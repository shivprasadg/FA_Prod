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
Dim ModeX      As String
Dim CheckUpdate As String
Dim CurrentRate As Double

Public Sub SetFormLevel()
    SourceX = "TaxStatesExceptions"
    TaxLevelIDx = "TaxStatesID"
    LevelX = "State"
    TargetIDx = 11
    CurrentRate = 0
    On Error GoTo FixMe
    If Me.Parent.Form.Name = "Tax_City" Then
        SourceX = "TaxCitiesExceptions"
        TaxLevelIDx = "TaxCitiesID"
        LevelX = "City"
        TargetIDx = Me.Parent.CityID
        CurrentRate = Nz(Me.Parent!Tax_CityBaseSub.Form!TaxNumber, 0)
    End If
    If Me.Parent.Form.Name = "Tax_State" Then
        SourceX = "TaxStatesExceptions"
        TaxLevelIDx = "TaxStatesID"
        LevelX = "State"
        TargetIDx = Me.Parent.StateID
        CurrentRate = Nz(Me.Parent!Tax_StateBaseSub.Form!TaxNumber, 0)
    End If
    If Me.Parent.Form.Name = "Tax_County" Then
        SourceX = "TaxCountiesExceptions"
        TaxLevelIDx = "TaxCountiesID"
        LevelX = "County"
        TargetIDx = Me.Parent.CountyID
        CurrentRate = Nz(Me.Parent!Tax_CountyBaseSub.Form!TaxNumber, 0)
    End If
FixMe:
End Sub



Private Sub ExceptionDescrip_Click()
Me.ExceptionDescrip.ControlTipText = Left(RemoveRichText(Nz(Me.TaxScript, "Description")), 252) & "..."
End Sub

Private Sub Form_Current()
    On Error GoTo FixMe
    Me.Parent.TaxTypeId = Me.TaxTypeId
    Me.txtCurrent = Nz(Me.ID, 0)
FixMe:
End Sub

Private Sub Form_Load()
'    On Error GoTo FixMe
    Call SetFormLevel
    SQLz = SQLz & "SELECT EXP.ID "
    SQLz = SQLz & ",EXP.ExceptionDescription "
    SQLz = SQLz & ",EXP.TaxRule "
    SQLz = SQLz & ",EXP.TaxRate "
    SQLz = SQLz & ",'W1#' & EXP.FormURL1 & '#' As Web1 "
    SQLz = SQLz & ",'W2#' & EXP.FormURL2 & '#' As Web2 "
    SQLz = SQLz & ",EXP.IncludedNAICS "
    SQLz = SQLz & ",EXP.AppliesToTractors "
    SQLz = SQLz & ",EXP.AppliesToTrailers "
    SQLz = SQLz & ",EXP.AppliesToStraightTrucks "
    SQLz = SQLz & ",EXP.WeightLimitMinLbs "
    SQLz = SQLz & ",EXP.WeightLimitMaxLbs "
    SQLz = SQLz & ",EXP.TermExceedsMonths "
    SQLz = SQLz & ",EXP.TaxScript "
    SQLz = SQLz & ",EXP.StartsOn "
    SQLz = SQLz & ",EXP.EndsOn "
    SQLz = SQLz & ",EXP." & TaxLevelIDx & ", EXP.TaxTypeId "
    SQLz = SQLz & ",TaxType.TaxFullName "
    SQLz = SQLz & ",IIf(EXP.StartsOn > date(),0,IIf(EXP.StartsOn <= date() And nz(EXP.EndsOn,cdate('1/1/2050')) >=date(),-1,0)) AS IsActive"
    SQLz = SQLz & ",TaxType.TaxAbbreviation "
    SQLz = SQLz & ",TaxType.TaxDescription "
    SQLz = SQLz & ",EXP.TaxMaxValue "
    SQLz = SQLz & ",EXP.IncludedNAICS "
    SQLz = SQLz & "FROM " & SourceX & " AS EXP LEFT JOIN TaxType ON EXP.TaxTypeID = TaxType.Id;"
    Me.Form.RecordSource = SQLz
FixMe:
End Sub

Private Sub TaxNumber_AfterUpdate()
    CheckUpdate = AskAboutTaxChange(MakePercent(Me.TaxNumber), MakePercent(CStr(CurrentRate)), Me.TaxNumber)
    If CheckUpdate = "Undo" Then
        Me.Undo
        Exit Sub
    End If
    If CheckUpdate = "Correction" Then
        Me.Undo
        Me.TaxNumber = MakePercent(Me.TaxNumber)
        ModeX = "Edit"
    End If
    If CheckUpdate = "AddTax" Then
        Me.Undo
        ModeX = "Add"
    End If
    Call SetFormLevel
    DoCmd.close acForm, "Tax_Add_Exception"
    DoCmd.OpenForm "Tax_Add_Exception", acNormal, , , , , "" & LevelX & "|" & Me.TaxTypeId & "|" & TargetIDx & "|" & Me.TaxNumber & "|" & ModeX
    '   If LevelX = "State" Then DoCmd.OpenForm "Tax_Add_Exception", acNormal, , , , , "State|" & Me.TaxTypeId & "|" & Me.Parent.StateID & "|" & Me.TaxNumber & "|" & ModeX
    '   If LevelX = "County" Then DoCmd.OpenForm "Tax_Add_Exception", acNormal, , , , , "County|" & Me.TaxTypeId & "|" & Me.Parent.CountyID & "|" & Me.TaxNumber & "|" & ModeX
    'End If
End Sub
Private Sub EditMe_Click()
    Call SetFormLevel
    ModeX = "Edit"
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TaxExceptionsTemp")
    SQLz = "INSERT INTO TaxExceptionsTemp "
    SQLz = SQLz & "(" & TaxLevelIDx & ""
    SQLz = SQLz & ",TaxTypeID"
    SQLz = SQLz & ",TaxExceptionTypeID"
    SQLz = SQLz & ",ExceptionDescription"
    SQLz = SQLz & ",TaxRule"
    SQLz = SQLz & ",TaxRate"
    SQLz = SQLz & ",FormURL1"
    SQLz = SQLz & ",FormURL2"
    SQLz = SQLz & ",IncludedNAICS"
    SQLz = SQLz & ",AppliesToTractors"
    SQLz = SQLz & ",AppliesToTrailers"
    SQLz = SQLz & ",AppliesToStraightTrucks"
    SQLz = SQLz & ",WeightLimitMinLbs"
    SQLz = SQLz & ",WeightLimitMaxLbs"
    SQLz = SQLz & ",TermExceedsMonths"
    SQLz = SQLz & ",TaxScript"
    SQLz = SQLz & ",StartsOn"
    SQLz = SQLz & ",EndsOn"
    SQLz = SQLz & ",TaxMaxValue"
    SQLz = SQLz & ",TaxLevel"
    SQLz = SQLz & ",Mode"
    SQLz = SQLz & ",TaxExceptionID"
    SQLz = SQLz & ") Select"
    SQLz = SQLz & " " & TaxLevelIDx & " "
    SQLz = SQLz & ",TaxTypeID"
    SQLz = SQLz & ",TaxExceptionTypeID"
    SQLz = SQLz & ",ExceptionDescription"
    SQLz = SQLz & ",TaxRule"
    SQLz = SQLz & ",TaxRate"
    SQLz = SQLz & ",FormURL1"
    SQLz = SQLz & ",FormURL2"
    SQLz = SQLz & ",IncludedNAICS"
    SQLz = SQLz & ",AppliesToTractors"
    SQLz = SQLz & ",AppliesToTrailers"
    SQLz = SQLz & ",AppliesToStraightTrucks"
    SQLz = SQLz & ",WeightLimitMinLbs"
    SQLz = SQLz & ",WeightLimitMaxLbs"
    SQLz = SQLz & ",TermExceedsMonths"
    SQLz = SQLz & ",TaxScript"
    SQLz = SQLz & ",StartsOn"
    SQLz = SQLz & ",EndsOn"
    SQLz = SQLz & ",TaxMaxValue"
    SQLz = SQLz & ",'" & LevelX & "' "
    SQLz = SQLz & ",'" & ModeX & "' "
    SQLz = SQLz & ", " & Me.ID & " "
    SQLz = SQLz & "From " & SourceX & " WHERE ID=" & Me.ID
    DoCmd.RunSQL (SQLz)
    Dim NAICSx As String
    Dim xx     As Integer
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TaxExceptionsNAICSTemp")
    NAICSx = Nz(DLookup("IncludedNAICS", "TaxExceptionsTemp"), "")
    If NAICSx <> "" Then
        If Left(NAICSx, 1) <> "|" Then NAICSx = "|" & NAICSx
        For xx = 1 To HowManyInStr(NAICSx, "|")
            SQLz = "Insert into TaxExceptionsNAICSTemp (NAICS, NAICSDesc, TaxLevel,TaxExceptionID) " & _
                   "Values('" & SplitString(NAICSx, "|", xx) & "','" & DLookup("NAICSDesc", "NAICScodes", "NAICS='" & SplitString(NAICSx, "|", xx) & "'") & "','" & LevelX & "'," & Me.ID & ")"
            DoCmd.RunSQL (SQLz)

        Next
    End If
    DoCmd.SetWarnings True
    DoCmd.close acForm, "Tax_Add_Exception"
    DoCmd.OpenForm "Tax_Add_Exception", acNormal, , , , , "" & LevelX & "|" & Me.TaxTypeId & "|" & TargetIDx & "|" & CurrentRate & "|" & ModeX
End Sub

Public Sub TaxNumber_GotFocus()
    CurrentRate = Me.TaxNumber
End Sub
