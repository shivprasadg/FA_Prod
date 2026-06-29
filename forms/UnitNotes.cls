Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST As DAO.Recordset

Private Sub AddUnitNote_Click()
 DoCmd.GotoRecord , , acNewRec
 Me.UnitRef = Forms!Unitsform.UnitID
 Me.uEmp = FAUserInit
 Me.uNote.SetFocus
End Sub


Private Sub Form_Close()
   On Error GoTo Form_Close_Error
   On Error GoTo 0
   Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_UnitNotes"
End Sub
Public Function LoadData(UnitIDx As Long)
    On Error Resume Next
    If Nz(Me.UnitRef, 0) = 0 Or Nz(Me.UnitRef, 0) <> UnitIDx Then
        Dim SQLz   As String
        On Error GoTo 0
        SQLz = "SELECT * FROM UnitNotes WHERE UnitRef = " & UnitIDx & " ORDER BY uNoteDate DESC;"
        Me.Form.RecordSource = SQLz
        SQLz = "SELECT uNoteID, uNoteDate AS [Date], UnitRef "
        SQLz = SQLz & "FROM UnitNotes "
        SQLz = SQLz & "WHERE UnitRef=" & UnitIDx & " ORDER BY uNoteDate DESC;"
        Me.NotesList.RowSource = SQLz
    End If
End Function

Private Sub Form_Dirty(Cancel As Integer)
    If IsNull(Me.UnitRef) Then Me.UnitRef = Forms!Forms!Unitsform.UnitID
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Me.Dirty And IsNull(Me.UnitRef) Then Me.UnitRef = Forms!Unitsform.Form.UnitID
End Sub
Private Sub NotesList_Click()
On Error GoTo Fxme
Me.Form.FilterOn = False
 Set RST = Me.RecordsetClone
       RST.FindFirst "uNoteID = " & Me.NotesList.Column(0) & ""
       Me.Bookmark = RST.Bookmark
    RST.close
    Set RST = Nothing
Exit Sub
Fxme:
MsgBox "No Matching Record", vbCritical, AtlaasVersion
RST.close
Set RST = Nothing
End Sub

Private Sub SaveMe_Click()
If Me.Dirty Then
Me.Dirty = False
Me.NotesList.Requery
End If
End Sub

Private Sub uNote_GotFocus()
Me.uNote.BackColor = RGB(255, 251, 171)
End Sub

Private Sub uNote_LostFocus()
Me.uNote.BackColor = RGB(255, 255, 255)
End Sub
