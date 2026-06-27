Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Private Sub Form_Current()
On Error GoTo FixMe
    Me.txtCurrent = Me.ID
    Parent.CountyID = Me.ID
FixMe:
End Sub

Public Sub GotoRecord(IDX As Integer)
 Dim rs As DAO.Recordset
  Set rs = Me.RecordsetClone
  rs.FindFirst "[ID]=" & IDX
     Me.Bookmark = rs.Bookmark
  Set rs = Nothing
End Sub
