Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
   On Error GoTo 0
End Sub

Public Function LoadData(UnitIDx As Long)
  On Error Resume Next
    If Nz(Me.UnitID, 0) = 0 Or Nz(Me.UnitID, 0) <> UnitIDx Then
     On Error GoTo 0
     Dim SQLx As String
     SQLx = SQLx & "SELECT Schedule.SchNo, SchGrp.SGrpNO, EO.UnitID, EO.UserName, EO.DatePosted, EO.Notes, XD.ReplacedByUnitID, EO.ReplacedByScheduleID "
     SQLx = SQLx & ", EO.ReplacedByGroupID, XD.StartOrPurposedDate, XD.EndOrActualDate, XD.FinalRentDate, XD.oGroupID, XD.nGroupID, XD.oBLCD, XD.BDay "
     SQLx = SQLx & ", XD.exStartDate, XD.exFinalRent, XD.exActual, XD.oProjEOL, XD.nEstDelv, XD.nActualDelv, XD.nAccepted, XD.nPDStart, XD.nInServ, XD.nOLDate "
     SQLx = SQLx & "FROM ((UnitExitOption as EO LEFT JOIN vw_EXchangeITDates as XD ON EO.UnitID = XD.UnitID) "
     SQLx = SQLx & "LEFT JOIN SchGrp ON EO.ReplacedByGroupID = SchGrp.SGrpID) "
     SQLx = SQLx & "LEFT JOIN Schedule ON EO.ReplacedByScheduleID = Schedule.SchID "
     SQLx = SQLx & "Where OptionTypeID = 1 "
     SQLx = SQLx & "AND  EO.UnitID=" & UnitIDx & ";"
     Me.Form.RecordSource = SQLx

   Call SetControls
    End If
    On Error GoTo 0
End Function
Private Sub SetControls()

On Error GoTo FixMe
    If Nz(Me.CurrentRecord, 0) <> 0 Then
        Me.ModifyMe.caption = "Edit Option"
        Me.Detail.Visible = True
    Else
        Me.ModifyMe.caption = "Add Option"
        Me.Detail.Visible = False
    End If
    If Me.Form.NewRecord = False Then

    End If
    Exit Sub
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form - Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
   Resume Next
End Sub

Private Sub Form_Current()
     If DLookup("LeasePaidAdvArr", "SchGrp", "sGrpID=" & Me.Parent.GroupRef) = "Arr" Then
        Me.cProJOffLease = DateAdd("d", -1, DateAdd("m", 0, [FinalRentDate]))
     Else
        Me.cProJOffLease = DateAdd("d", -1, DateAdd("m", 1, [FinalRentDate]))
     End If
End Sub

Private Sub ModifyMe_Click()
If IsNull(Me.Parent.UnitID) Then Exit Sub
Call MakeExitOptionUnit(Me.Parent.UnitID, Me.Parent.SchID, Me.Parent.GroupRef, "EXchangeIT")
End Sub
