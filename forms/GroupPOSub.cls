Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Sub checkPO()
If Me.PONum1 & "test" = "test" Then
        Dim NextSeq As Integer
        NextSeq = CInt(DMax("PoSeq", "Last10PONumbers")) + 1
        Dim ResposeX As VbMsgBoxResultEx
        Dim msg1 As String
        msg1 = msg1 & "<font size=3>Asign a PO Number</font><br />The Next Valid PO Number is:"
        msg1 = msg1 & "<br/><b>" & Me.SGrpFAID & "-" & NextSeq & "</b>"
        msg1 = msg1 & "<br /> would you like to assign it to this group?<br />"
        msg1 = msg1 & "   Please select..."
        ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
        If ResposeX = vbBt2 Then Exit Sub
        If ResposeX = vbBt1 Then
            Me.PONum1 = Me.SGrpFAID & "-" & NextSeq
            If Me.Form.Dirty Then Me.Form.Dirty = False
        End If
    End If
End Sub
Private Sub MakePO_Click()
    Call checkPO
    If Me.PONum1 & "test" = "test" Then
        MsgBox "Sorry but you haven't assigned a PO number" & vbNewLine & "The reports will not run", vbCritical, AtlaasVersion
        Exit Sub
    End If
    DoCmd.close acReport, "PO_Vendor"
    DoCmd.OpenReport "PO_Vendor", acViewPreview
End Sub

Private Sub MakePOTracker_Click()
    Call checkPO
    If Me.PONum1 & "test" = "test" Then
        MsgBox "Sorry but you haven't assigned a PO number" & vbNewLine & "The reports will not run", vbCritical, AtlaasVersion
        Exit Sub
    End If
    DoCmd.close acReport, "PO_VendorTracking"
    DoCmd.OpenReport "PO_VendorTracking", acViewPreview
End Sub

Private Sub PONum1_Click()
Me.Parent.ActivePO = Me.PONum1
End Sub

Private Sub POSentDt_Click()
Me.Parent.ActivePO = Me.PONum1
End Sub

Private Sub POVendor_Click()
Me.Parent.ActivePO = Me.PONum1
End Sub
Private Sub PONum2_Click()
Me.Parent.ActivePO = Me.PONum2
End Sub

Private Sub PO2SentDt_Click()
Me.Parent.ActivePO = Me.PONum2
End Sub

Private Sub PO2Vendor_Click()
Me.Parent.ActivePO = Me.PONum2
End Sub
Private Sub PONum3_Click()
Me.Parent.ActivePO = Me.PONum3
End Sub

Private Sub PO3SentDt_Click()
Me.Parent.ActivePO = Me.PONum3
End Sub
Private Sub PO3Vendor_Click()
Me.Parent.ActivePO = Me.PONum3
End Sub
Private Sub PONum4_Click()
Me.Parent.ActivePO = Me.PONum4
End Sub

Private Sub PO4SentDt_Click()
Me.Parent.ActivePO = Me.PONum4
End Sub
Private Sub PO4Vendor_Click()
Me.Parent.ActivePO = Me.PONum4
End Sub
