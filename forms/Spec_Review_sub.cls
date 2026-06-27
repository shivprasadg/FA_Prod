Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function DiscountedPrice() As String

Dim value1 As Long
Dim value2 As Long
Dim value3 As Long
Dim SubtotalPrice As Long
Dim discount As Long

If IsNull(SRSellingPrice) Then
    value1 = 0
Else
    value1 = SRSellingPrice
End If

If IsNull(SRListPrice) Then
    value2 = 0
Else
    value2 = SRListPrice
End If

If IsNull(SROtherPrice) Then
    value3 = 0
Else
    value3 = SROtherPrice
End If

'Set the DiscountedPrice value
SubtotalPrice = (value1 / value2) + value3
discount = (value1 / value2) * 0.01
MsgBox discount
DiscountedPrice = SubtotalPrice - discount

End Function

Public Function TotalRebate() As Integer
'set values to zero if null

If IsNull(SRWBRebate) Then
    SRWBRebate = 0
End If

If IsNull(SREngineRebate) Then
    SREngineRebate = 0
End If

If IsNull(SRTransmissionRebate) Then
    SRTransmissionRebate = 0
End If

If IsNull(SRFrontAxleRebate) Then
    SRFrontAxleRebate = 0
End If

If IsNull(SRRearAxleRebate) Then
    SRRearAxleRebate = 0
End If

If IsNull(SRRatioRebate) Then
    SRRatioRebate = 0
End If

If IsNull(SRSuspensionRebate) Then
    SRSuspensionRebate = 0
End If

If IsNull(SR5thWheelRebate) Then
    SR5thWheelRebate = 0
End If

If IsNull(SRFrontWheelsRebate) Then
    SRFrontWheelsRebate = 0
End If

If IsNull(SRFrontTiresRebate) Then
    SRFrontTiresRebate = 0
End If

If IsNull(SRRearWheelsRebate) Then
    SRRearWheelsRebate = 0
End If

If IsNull(SRRearTiresRebate) Then
    SRRearTiresRebate = 0
End If

If IsNull(SRJakeRebate) Then
   SRJakeRebate = 0
End If

If IsNull(SRMfgRebate) Then
   SRMfgRebate = 0
End If

If IsNull(SRMiscRebate1) Then
   SRMiscRebate1 = 0
End If

If IsNull(SRMiscRebate2) Then
   SRMiscRebate2 = 0
End If

If IsNull(SRMiscRebate3) Then
   SRMiscRebate3 = 0
End If

If IsNull(SRFilterRebate) Then
   SRFilterRebate = 0
End If

If IsNull(SRFanClutchRebate) Then
   SRFanClutchRebate = 0
End If

If IsNull(TrailerMfgRebate) Then
   TrailerMfgRebate = 0
End If

'get  the TotalRebate value
TotalRebate = SRWBRebate + SREngineRebate + SRTransmissionRebate + SRFrontAxleRebate + SRRearAxleRebate + SRRatioRebate + SRSuspensionRebate + SR5thWheelRebate + SRFrontWheelsRebate + SRFrontTiresRebate + SRRearWheelsRebate + SRRearTiresRebate + SRJakeRebate + SRMfgRebate + SRMiscRebate1 + SRMiscRebate2 + SRMiscRebate3 + SRFanClutchRebate + SRFilterRebate + TrailerMfgRebate



End Function

Private Sub CloseBtn_Click()
On Error GoTo Err_CloseBtn_Click


    If Me.Dirty Then Me.Dirty = False
    DoCmd.close

Exit_CloseBtn_Click:
    Exit Sub

Err_CloseBtn_Click:
    MsgBox Err.Description
    Resume Exit_CloseBtn_Click

End Sub


Private Sub AddDocBtn_Click()

    Dim StrFilter As String
    Dim strInputFileName As String


Me.SREqSpecLink.Value = ahtCommonFileOpenSave( _
                Filter:=StrFilter, OpenFile:=True, _
                DialogTitle:="Please select an input file...", _
                flags:=ahtOFN_HIDEREADONLY)

End Sub

Private Sub AddQuoteBtn_Click()

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Quote_edit"

    DoCmd.OpenForm stDocName, , , , acFormAdd
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)

    If SRID > 0 Then
        Me.FormFooter.Visible = True And SaveSpecBtn.Visible = False
        Else: Me.FormFooter.Visible = False And SaveSpecBtn.Visible = True
    End If

    If SRID > 0 Then
      Me.Detail.Visible = True
      Else: Me.Detail.Visible = False
    End If

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

    If IsNull(SRID) = False Then
        Me.FormFooter.Visible = True
    End If


    If IsNull(SRID) = False Then
      Me.Detail.Visible = True
    End If

    If IsNull(SRID) = True Then
        SaveSpecBtn.Visible = True
    End If

End Sub

Private Sub Form_GotFocus()

    Forms!Spec_Review_sub.Quotes_sub.Requery

End Sub

Private Sub Form_Load()

    If IsNull(SRID) = False Then
        Me.FormFooter.Visible = True
    End If


    If IsNull(SRID) = False Then
      Me.Detail.Visible = True
    End If

    If IsNull(SRID) = True Then
        SaveSpecBtn.Visible = True
    End If

End Sub

Private Sub SaveSpecBtn_Click()
    SRNotes = ""
    DoCmd.Save
    DoCmd.Requery

   Me.FormFooter.Visible = True
   Me.Detail.Visible = True
   SaveSpecBtn.Visible = False

End Sub

Private Sub SR5thWheelRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRDiscountPrice_Click()
  SRDiscountPrice = DiscountedPrice
End Sub

Private Sub SREngineRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SREqSpecLink_Click()
    If SREqSpecLink <> "" Then
        Application.FollowHyperlink SREqSpecLink
        DoCmd.SetWarnings False
     End If
End Sub

Private Sub SRFanClutchRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRFilterRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRFrontAxleRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRFrontTiresRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRFrontWheelsRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRJakeRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRListPrice_AfterUpdate()
 SRDiscountPrice = DiscountedPrice
End Sub

Private Sub SRMfgRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRMiscRebate1_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRMiscRebate2_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRMiscRebate3_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SROtherPrice_AfterUpdate()
 SRDiscountPrice = DiscountedPrice
End Sub

Private Sub SRRatioRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRRearAxleRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRRearTiresRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRRearWheelsRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub


Private Sub SRRebateTotal_Click()

    SRRebateTotal = TotalRebate

End Sub

Private Sub SRSellingPrice_AfterUpdate()
 SRDiscountPrice = DiscountedPrice
End Sub

Private Sub SRSuspensionRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRTransmissionRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub SRWBRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub

Private Sub TrailerMfgRebate_AfterUpdate()
    SRRebateTotal = TotalRebate
End Sub
