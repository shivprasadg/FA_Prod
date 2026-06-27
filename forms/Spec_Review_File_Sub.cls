Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddDocBtn_Click()
On Error GoTo Err_AddDocBtn_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Attach_Spec_Rev"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_AddDocBtn_Click:
    Exit Sub

Err_AddDocBtn_Click:
    MsgBox Err.Description
    Resume Exit_AddDocBtn_Click

End Sub

Private Sub Form_Load()

'Forms!Spec_Review_File_Sub.RecordsetClone.MoveLast

Me.RecordsetClone.MoveLast

MsgBox "My form contains " _
    '& Forms!Spec_Review_File_Sub.RecordsetClone.RecordCount
    Me.RecordsetClone.RecordCount
    '& " records.", vbInformation, "Record Count"

End Sub


Private Sub OpenFileBtn_Click()
    Application.FollowHyperlink AttachedFilePath

    DoCmd.SetWarnings False

End Sub

Private Sub RemoveFileBtn_Click()

    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDeleteRecord
End Sub
