Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub AmendmentsList_Click()
Dim RST As DAO.Recordset
    Set RST = Me.RecordsetClone
    RST.FindFirst "ID = " & Me.AmendmentsList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
End Sub

Private Sub Form_Current()
'Dim SQLz As String
'SQLz = "SELECT mlamends.ID, mlamends.MLAmtNO AS AmendNum, mlamends.MLAmdDt AS AmendDate, mlamends.MLRef " & _
       "From mlamends " & _
       "WHERE MLRef =" & Me.Parent.MLID & " " & _
       "ORDER BY MLAmtNO DESC , MLAmdDt DESC;"
 '     Me.AmendmentsList.RowSource = SQLz
End Sub
