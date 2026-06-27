Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Private Sub Form_Current()
    On Error GoTo FixMe
    Me.txtCurrent = Me.ID
  '  Me.Parent.StateID = Me.ID
  '  SQLz = "Select * From TaxCounties WHERE TaxStatesID=" & Me.ID & " Order By CountyName"
  '  Forms(Me.Parent.Form.Name)!Tax_CountySub.Form.RecordSource = SQLz
  '  Forms(Me.Parent.Form.Name)!Tax_CountySub.Form.Detail.Visible = True
FixMe:
End Sub

Public Sub GotoRecord(IDX As Integer)
 Dim rs As DAO.Recordset
  Set rs = Me.RecordsetClone
  rs.FindFirst "[ID]=" & IDX
     Me.Bookmark = rs.Bookmark
  Set rs = Nothing
End Sub
