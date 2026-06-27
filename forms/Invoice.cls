Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub InvoicedUnitType_Click()
If Me.InvoicedUnitType = "Child Part" Then
SQLz = "SELECT ChildVendorID, FAID, MLID, clientID, CompanyName " & _
       "From vw_Sch_Grp_Units_ChildParts  " & _
       "GROUP BY ChildVendorID, FAID, MLID, clientID, CompanyName " & _
       "HAVING ChildVendorID=" & Me.VendorID & ";"
Else
SQLz = "SELECT VendorId, FAID, MLID, clientID, CompanyName " & _
       "FROM vw_Sch_Grp_Units_Leased " & _
       "GROUP BY VendorId, FAID, MLID, clientID, CompanyName " & _
       "Having VendorId=" & Me.VendorID & " ;"
End If
Me.FAIDCbo.RowSource = SQLz
End Sub

Private Sub OpenClientView_Click()
DoCmd.OpenForm "Search_Clients"
End Sub

Private Sub OpenUnitLookUp_Click()
DoCmd.OpenForm "UnitReverseLookUpForm"
End Sub
