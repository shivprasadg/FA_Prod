Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub EditSched_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Deals_view"

    stLinkCriteria = "[dealID]=" & Me![dealID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub


Private Sub RefreshBtn_Click()
On Error GoTo Err_RefreshBtn_Click

    Dim stDocName As String

    stDocName = "Requery"
    DoCmd.RunMacro stDocName

Exit_RefreshBtn_Click:
    Exit Sub

Err_RefreshBtn_Click:
    MsgBox Err.Description
    Resume Exit_RefreshBtn_Click

End Sub

Private Sub BnkSchBtn_Click()
    DoCmd.close acForm, "Schedule_Form"
    DoCmd.OpenForm "Schedule_Form", , , "SchID=" & Me.SchID
End Sub
