Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
    If Me.Dirty Then Me.Dirty = False
End Sub

Private Sub Form_Current()
    If Me.UnitVINLocked = -1 Then
        Me.UnitMSORecd.Locked = True
        Me.UnitMSOCopy.Locked = True
    Else
        Me.UnitMSORecd.Locked = False
        Me.UnitMSOCopy.Locked = False
    End If
End Sub


Private Sub UnitMSOCopy_AfterUpdate()
Call MarkVIN
End Sub

Private Sub UnitMSOCopy_GotFocus()
If Not IsDate(Me.UnitMSORecd) And Not IsDate(Me.UnitMSOCopy) Then Me.UnitMSOCopy = Date
Call MarkVIN
End Sub

Private Sub UnitMSORecd_AfterUpdate()
Call MarkVIN
End Sub

Private Sub UnitMSORecd_GotFocus()
If Not IsDate(Me.UnitMSORecd) And Not IsDate(Me.UnitMSOCopy) Then Me.UnitMSORecd = Date
Call MarkVIN
End Sub

Private Sub UnitVINLocked_Click()
Call MarkVIN
End Sub

Private Sub MarkVIN()
    If IsDate(Me.UnitMSORecd) Or IsDate(UnitMSOCopy) Then
        UnitVINLocked = True
        If Not IsDate(UnitVINVer) Then Me.UnitVINVer = Date
        If Me.Form.Dirty Then Me.Form.Dirty = False
    Else
    Me.UnitVINVer = ""
        If UnitVINLocked = True Then
            MsgBox "Before a VIN can be Verified & Locked" & vbNewLine & "" & _
                   "A Unit Must have either:" & vbNewLine & "" & _
                   "an MSO or an MSO Copy logged with a proper date", vbCritical, AtlaasVersion
            UnitVINLocked = False
            If Me.Form.Dirty Then Me.Form.Dirty = False
        End If
    End If
End Sub

Private Sub UnitVINVer_AfterUpdate()
Call MarkVIN
End Sub
