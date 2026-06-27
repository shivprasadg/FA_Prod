Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

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

Private Sub Form_Load()
If Not IsNull(pricingSynSpread) And Not IsNull(pricingSynPerc) Then

    pricingSynYield = pricingSynSpread + pricingSynPerc

Else
    pricingSynYield = 0
End If



End Sub

Private Sub pricingSpread_AfterUpdate()
If Not IsNull(pricingSynSpread) And Not IsNull(pricingSynPerc) Then

    pricingSynYield = pricingSynSpread + pricingSynPerc
Else
    pricingSynYield = 0

End If

End Sub

Private Sub pricingSyndBank_DblClick(Cancel As Integer)

    DoCmd.OpenForm "Bank_add", acDesign, , , acFormAdd, acDialog

End Sub

Private Sub pricingSynPerc_AfterUpdate()
If Not IsNull(pricingSynSpread) And Not IsNull(pricingSynPerc) Then

    pricingSynYield = pricingSynSpread + pricingSynPerc

Else
    pricingSynYield = 0
End If

End Sub

Private Sub pricingYield_Click()

If Not IsNull(pricingSynSpread) And Not IsNull(pricingSynPerc) Then

    pricingSynYield = pricingSynSpread + pricingSynPerc

Else
    pricingSynYield = 0
End If

End Sub
