VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFieldEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' INTERNAL STATE
' ============================================================
Private PrevValue      As Variant
Private pField         As String
Private pLabel         As String
Private pLevel         As String
Private pUnitSpecID    As Long
Private pGroupSpecID   As Long

' ============================================================
' PUBLIC API — CALLED FROM THE FORM
' ============================================================

' Called from Form_Current (every time record changes)
Public Sub SetSpecIDs(ByVal UnitSpecID As Long, ByVal GroupSpecID As Long)
    pUnitSpecID = UnitSpecID
    pGroupSpecID = GroupSpecID
End Sub

' Called by: Control_GotFocus
Public Sub HandleGotFocus(ctl As Control)
    LoadMeta ctl
    PrevValue = ctl.Value
End Sub

' Called by: Control_AfterUpdate
Public Sub HandleAfterUpdate(ctl As Control)
    Dim SpecId As Long
    Dim ok As Boolean

    If pField = "" Then Exit Sub

    ' Pick correct SpecID based on level
    Select Case pLevel
        Case "Unit":  SpecId = pUnitSpecID
        Case "Group": SpecId = pGroupSpecID
        Case Else:    Exit Sub
    End Select

    If Nz(PrevValue, "") <> Nz(ctl.Value, "") Then

        ok = ChangeItems(ctl.Value, PrevValue, pField, pLabel, pLevel, SpecId)

        If Not ok Then
            ' User declined — undo the change
            ctl.Undo
            PrevValue = Null
            pField = ""
            ctl.SetFocus
            MsgBox "The " & pLabel & " change was undone.", vbInformation, "It Never Happened"
            Exit Sub
        End If

    End If
End Sub

' Called by: Control_DblClick
Public Sub HandleDblClick(ctl As Control, Cancel As Integer)
    Dim SpecId As Long
    Dim ok As Boolean

    LoadMeta ctl
    If pField = "" Then Exit Sub

    ' Pick correct SpecID based on level
    Select Case pLevel
        Case "Unit":  SpecId = pUnitSpecID
        Case "Group": SpecId = pGroupSpecID
        Case Else:    Exit Sub
    End Select

    If Trim(Nz(ctl.Value, "")) <> "" Then

        ok = SetItems(ctl.Value, pField, pLabel, pLevel, SpecId)

        If Not ok Then
            ' User declined — nothing to undo (SetItems does not change the control)
            MsgBox "The " & pLabel & " was not updated.", vbInformation, "It Never Happened"
            Exit Sub
        End If

    End If
End Sub

' ============================================================
' METADATA LOOKUP
' ============================================================
Private Sub LoadMeta(ctl As Control)
    Select Case ctl.Name

        Case "PowerType"
            pField = "PowerType"
            pLabel = "Power Type"
            pLevel = "Unit"

        Case "UnitsSRent"
            pField = "UnitsSRent"
            pLabel = "Sch Rent"
            pLevel = "Unit"

        Case "GroupName"
            pField = "GroupName"
            pLabel = "Group Name"
            pLevel = "Group"

        Case "UnitType"
            pField = "UnitType"
            pLabel = "Unit Type"
            pLevel = "Unit"

        Case "UnitUOM"
            pField = "UnitUOM"
            pLabel = "Unit UOM"
            pLevel = "Unit"

        Case Else
            pField = ""
            pLabel = ""
            pLevel = ""
    End Select
End Sub
