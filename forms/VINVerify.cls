Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz As String

Private Sub CloseMe_Click()
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub LessThan17_Click()
Me.seekvin = ""
Me.VIN_List.RowSource = ""
End Sub

Private Sub SeekVin_Change()
If Me.LessThan17 = False Then
   If Len(Me.seekvin.Text) < 17 Then Exit Sub
Exit Sub
End If
'SQLz = "SELECT Units.UnitID AS AssetID, Units.UnitVIN " & _
       "From Units WHERE UnitVIN like '*" & Me.seekvin.Text & "*' " & _
       "ORDER BY Units.UnitVIN;"
SQLz = "SELECT Units.UnitID, Units.UnitVIN AS AssetID " & _
       "From Units WHERE (UnitStatus = 'B' Or UnitStatus = 'A' or UnitStatus = 'TR') AND UnitVIN like '*" & Me.seekvin.Text & "*' " & _
       "ORDER BY Units.UnitVIN;"
Me.VIN_List.RowSource = SQLz
End Sub
