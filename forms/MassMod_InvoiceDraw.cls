Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Dim ModMade As Boolean
Dim PrevDraw, Askit As Long

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    If ModMade = -1 Then Call CheckATFChanges(Me.groupID, "Rent Change")
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub DrawDownId_Change()
    Askit = 0
    If Nz(PrevDraw, 0) <> Me.DrawDownId Then
        Dim CPCost As Currency
        CPCost = Nz(DSum("childFinalCost", "ChildParts", "UnitREf=" & Me.UnitID), 0)
        If CPCost > 0 Then
            Askit = MsgBox("Include this units Child Parts," & vbNewLine & "and Sync the Invoice Number?", vbQuestion + vbYesNo + vbCritical, AtlaasVersion & " Do Task?")
        End If
        'iBankID = DLookup("DrawBankID", "DrawDowns", "DrawDownID=" & Me.DrawDownId)
        DoCmd.SetWarnings False
        SQLz = "Update Units Set UnitIFBankID=" & Me.DrawDownId.Column(4) & " ,UnitIFAmt = VendorCost "
        If Askit = 6 Then
            SQLz = SQLz & "+" & CPCost & " "
            SQLy = "UPDATE ChildParts Set DrawDownId=" & Me.DrawDownId & " "
            SQLy = SQLy & ", ChildInvoiceNum='" & Me.UnitInvoiceNum & "' "
            SQLy = SQLy & "WHERE UnitRef=" & Me.UnitID & ";"
            DoCmd.RunSQL (SQLy)
        End If
        SQLz = SQLz & "WHERE UnitID =" & Me.UnitID
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.Form.Requery
    End If
End Sub

Private Sub DrawDownId_GotFocus()
  PrevDraw = Nz(Me.DrawDownId, 0)
End Sub

Private Sub Form_Current()
    On Error GoTo whoops
    Me.txtCurrent = Me.UnitID
whoops:
    Exit Sub
End Sub

Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub Form_Load()
If Nz(Me.OpenArgs, "") <> "" Then
 Me.Form.RecordSource = "Select * From Units Where GroupRef=" & SplitString(Me.OpenArgs, "?", 0) & ";"
 Me.SchNo = SplitString(Me.OpenArgs, "?", 1)
End If
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   ModMade = -1
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Me.UnitVIN.SetFocus
End Sub

Private Sub DrawDownSubmit_Click()
    On Error GoTo FixMe
    Dim ExtraFields As String
    If Nz(Me.DrawDownCbo, 1) = 1 Then
        Dim Committ As Variant
        Committ = MsgBox("You're about to Remove Units From a Draw" & vbNewLine & "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Remove Draw!")
        If Committ = vbNo Then Exit Sub
    End If
    If Nz(Me.DrawDownCbo, 1) > 1 Then
       ExtraFields = ", Units.UnitIFBankID= " & Me.DrawDownId.Column(4) & " , Units.UnitIFAmt = VendorCost "
    Else
       ExtraFields = ", Units.UnitIFBankID=Null , Units.UnitIFAmt = 0 "
    End If
     Call UpdateMassUnits("Units.DrawDownId", Me.DrawDownCbo, , ExtraFields)
FixMe:
    Me.Form.Refresh
End Sub
Private Sub VendorPayDateSubmit_Click()
    On Error GoTo FixMe
    If IsDate(VendorPayDate) Then Call UpdateMassUnits("Units.UnitVendorPytDate", Me.VendorPayDate)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub UnitInvoiceNumSubmit_Click()
    On Error GoTo FixMe
    If UnitInvoiceNumCBO & "TEST" <> "TEST" Then Call UpdateMassUnits("Units.UnitInvoiceNum", Me.UnitInvoiceNumCBO)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub InvoiceDateSubmit_Click()
    If IsDate(UnitInvoiceDateCBO) Then Call UpdateMassUnits("Units.UnitInvoiceDate", Me.UnitInvoiceDateCBO)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
