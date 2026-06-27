Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST As DAO.Recordset
Dim SQLz As String

Private Sub AddNew_Click()

Dim HoldAssetX As String
If IsNull(Me.AssetID) Then HoldAssetX = Me.Parent.Units_List.Column(0) Else HoldAssetX = Me.AssetID
DoCmd.GotoRecord , , acNewRec
Me.AssetID = HoldAssetX
End Sub

Private Sub ChangeMe_Click()
Me.VendorID.Locked = False
Me.ExpenseTypeId.Locked = False
 Me.EditHereAlso.Visible = True
'DoCmd.close acForm, "RemarketExpense_Edit"
'DoCmd.OpenForm "RemarketExpense_Edit", , , "RMKTWarrantyId =" & Me.RMKTWarrantyId
End Sub



Private Sub DeleteMe_Click()
 Dim Response As String
    Response = MsgBox("Are you sure you want to Delete this Expense?", vbQuestion + vbYesNo + vbCritical, "MDK!!")
    If Response = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunCommand acCmdDeleteRecord
        DoCmd.SetWarnings True
        Me.Form.Requery
        Me.ExpenseList.Requery
    End If
End Sub

Private Sub Form_Close()
'DoCmd.OpenForm " FAMainMenu"
End Sub

Private Sub Form_Current()
If Me.Form.NewRecord Then
   Me.ExpenseTypeId.Locked = False
   Me.VendorID.Locked = False
   Me.EditHereAlso.Visible = True
Else
Me.ExpenseTypeId.Locked = True
Me.VendorID.Locked = True
Me.EditHereAlso.Visible = False
End If
End Sub

Private Sub Form_Load()
'SQLz = "SELECT RMKTExpenseId, AssetId, RemarketExpense AS Type, VendorSNm, ExpenseDate, Expense, InvoiceNumber, " & _
       "InvoiceDate, InvoiceAmount, RemarketExpense.InvoicePaidDate, RemarketExpense.Comment " & _
       "FROM (RemarketExpense LEFT JOIN Vendors ON RemarketExpense.VendorId = Vendors.VendorID) " & _
       "LEFT JOIN RemarketExpenseType ON RemarketExpense.ExpenseTypeId = RemarketExpenseType.RemarketExpenseTypeId " & _
       "WHERE AssetId = " & Nz(Me.AssetID, 0) & ";"
'       Me.ExpenseList = SQLz
End Sub

Private Sub GoHome_Click()
DoCmd.OpenForm " FAMainMenu"
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub
Private Sub ExpenseList_Click()
   Set RST = Me.RecordsetClone
       RST.FindFirst "RMKTExpenseId = " & Me.ExpenseList.Column(0)
       Me.Bookmark = RST.Bookmark
   Set RST = Nothing
End Sub

Private Sub SaveUpdate_Click()
Me.Form.Refresh
Call Form_Current
End Sub
