Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Printme_Click()
DoCmd.OpenReport "RemarketingBillOfSale"
DoCmd.close acReport, "RemarketingBillOfSale"
End Sub
