Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddPricingBtn_Click()
On Error GoTo Err_AddPricingBtn_Click

    Dim stDocName As String

    stDocName = "Pricing_Add"
    DoCmd.OpenForm stDocName, , , , acFormAdd


Exit_AddPricingBtn_Click:
    Exit Sub

Err_AddPricingBtn_Click:
    MsgBox Err.Description
    Resume Exit_AddPricingBtn_Click

End Sub

Private Sub groupCost_AfterUpdate()

If groupUnits > 0 And GroupCost > 0 Then
groupTotal = groupUnits * GroupCost
Else: groupTotal = 0
End If

End Sub

Private Sub groupCost_Change()

If groupUnits > 0 And GroupCost > 0 Then
groupTotal = groupUnits * GroupCost
Else: groupTotal = 0
End If

End Sub

Private Sub AddCreditRequestBtn_Click()

    Dim strMessage As String
    Dim dbsUnits As DataBase
    Dim rstTypes As DAO.Recordset

'Open the Credit request table and set the GroupRef value.
Set dbsUnits = CurrentDb
Set rstTypes = dbsUnits.OpenRecordset("CreditRequests", dbOpenDynaset, dbSeeChanges)
rstTypes.AddNew
rstTypes!GroupRef = Me![groupID]
rstTypes.Update
MsgBox "The new Credit Request has been added", vbOKOnly
Forms!Groups_view.Credit_Request_sub.Requery

End Sub

Private Sub AddAUnitBtn_Click()
    Dim strMessage As String
    Dim dbsUnits As DataBase
    Dim rstTypes As DAO.Recordset

'Open the Units table and set the GroupRef value.
Set dbsUnits = CurrentDb
Set rstTypes = dbsUnits.OpenRecordset("Units", dbOpenDynaset, dbSeeChanges)
rstTypes.AddNew
rstTypes!GroupRef = Me![groupID]
rstTypes.Update
MsgBox "The new unit has been added", vbOKOnly
Forms!Groups_view.Units_sub.Requery

End Sub

Private Sub AddExpenseBtn_Click()
    Dim strMessage As String
    Dim dbsUnits As DataBase
    Dim rstTypes As DAO.Recordset

'Open the Expenses table and set the GroupRef value.
Set dbsUnits = CurrentDb
Set rstTypes = dbsUnits.OpenRecordset("Expenses", dbOpenDynaset, dbSeeChanges)
rstTypes.AddNew
rstTypes!GroupsRef = Me![groupID]
rstTypes!DealsRef = Me![dealRef]
rstTypes.Update
MsgBox "A blank expense has been added", vbOKOnly
Forms!Groups_view.Expenses_sub.Requery
End Sub

Private Sub CloseformBtn_Click()
If IsOpen("Deals_view") Then
    Forms!deals_view.Group_Summary_sub.Requery
    DoCmd.close

Else
    DoCmd.close

End If

End Sub

Private Sub Command103_Click()

Dim strForm
Dim strCriteria

strForm = "Units"
strCriteria = "[groupRef]=" & Forms!Groups_view.groupID

DoCmd.OpenForm strForm, , , strCriteria, , acDialog
End Sub

Private Sub Form_Current()
'DoCmd.Maximize
End Sub


Private Sub Form_Load()
'DoCmd.Maximize

End Sub

Private Sub groupBillingRef_Change()

Forms!Groups_view.Locations_Billing_sub.Requery

End Sub

Private Sub groupMakeRef_DblClick(Cancel As Integer)
    DoCmd.OpenForm "Make_add", , , , acFormAdd
End Sub



Private Sub groupModelRef_DblClick(Cancel As Integer)
    DoCmd.OpenForm "Model_add", , , , acFormAdd
End Sub



Private Sub groupTypeRef_DblClick(Cancel As Integer)
    DoCmd.OpenForm "groupType_add", , , , acFormAdd
End Sub

Private Sub groupUnits_AfterUpdate()

If groupUnits > 0 And GroupCost > 0 Then
groupTotal = groupUnits * GroupCost
Else: groupTotal = 0
End If

End Sub

Private Sub groupUnits_Change()

If groupUnits > 0 And GroupCost > 0 Then
groupTotal = groupUnits * GroupCost
Else: groupTotal = 0
End If


End Sub

Private Sub OpenUnitsSubbtn_Click()
On Error GoTo Err_OpenUnitsSubbtn_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Units"

    stLinkCriteria = "[GroupREF]=" & Me![groupID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OpenUnitsSubbtn_Click:
    Exit Sub

Err_OpenUnitsSubbtn_Click:
    MsgBox Err.Description
    Resume Exit_OpenUnitsSubbtn_Click

End Sub

Private Sub Lease_Pricing_Click()

    Forms!Groups_view.QuoteFinal_sub.Requery
End Sub

Private Sub Units_Click()
    DoCmd.Requery Forms!Groups_view.Units_sub
End Sub



'------------------------------------------------------------
' OpenLocationsView_btn_Click
'
'------------------------------------------------------------
Private Sub OpenLocationsView_btn_Click()
On Error GoTo OpenLocationsView_btn_Click_Err

Dim StrFormName
Dim StrFilter

StrFormName = "Locations_view"
StrFilter = "[clientRef]=" & Me.clientRef

    DoCmd.OpenForm StrFormName, acNormal, , StrFilter, , acNormal


OpenLocationsView_btn_Click_Exit:
    Exit Sub

OpenLocationsView_btn_Click_Err:
    MsgBox Error$
    Resume OpenLocationsView_btn_Click_Exit

End Sub
