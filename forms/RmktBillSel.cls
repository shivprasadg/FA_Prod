Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub BillInfoBtn_Click()
    DoCmd.close acForm, "ReMarketInvoiceManager"
    DoCmd.OpenForm "ReMarketInvoiceManager"
End Sub

Private Sub BillLog_Click()
    DoCmd.close acReport, "RemarketingBillLog"
    DoCmd.OpenReport "RemarketingBillLog", acViewReport
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub ExportLog2Excel_Click()
    FormatExcel "RmktBilllogq", "C:\Faas\" & FaUserName & "\Temp\RmktBilllogq.xlsx", "BillLog"

End Sub

Private Sub Form_Load()
    If DLookup("[Access Level]", "TBL_Employees", "[User ID] = FAUserName()") = 5 Then
        Me.BillInfoBtn.Visible = False
        Me.[10bill].Visible = False
        Me.[18bill].Visible = False
    Else
        Me.BillInfoBtn.Visible = True
    End If
End Sub
