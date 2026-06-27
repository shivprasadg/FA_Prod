Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private ranCurrent As Boolean
Private Sub AddFileBtn_Click()

    Dim SQL As String

    SQL = "INSERT INTO AttachedFiles ( AttachedFilePath, ClientREF, DealREF, GroupREF, UnitREF, PricingREF, AttachedFileTypeREF, AttachedFileDescShort, AttachedFileDescLong )" & _
          "SELECT [txtfilename] AS AttachedFilePath, ClientREF, DealREF, GroupREF, UnitREF, PricingREF, AttachedFileTypeREF, AttachedFileDescShort, AttachedFileDescLong "

    DoCmd.SetWarnings False

    DoCmd.RunSQL SQL

    DoCmd.SetWarnings True

    DoCmd.close

    If IsOpen("ClientsView") Then
        Forms!ClientsView.Attached_Files_sub.Requery
    ElseIf IsOpen("Deals_View") Then
        Forms!deals_view.Attached_Files_sub.Requery
    End If


End Sub

Private Sub AttachedFileTypeRef_DblClick(Cancel As Integer)

    Dim StrFormName As String

    StrFormName = "fileType_add"

    DoCmd.OpenForm StrFormName, acNormal, , , acFormAdd, acDialog

End Sub

Private Sub bBrowse_Click()

Dim StrFilter As String

Dim strInputFileName As String


Me.txtfilename.Value = ahtCommonFileOpenSave( _
                Filter:=StrFilter, OpenFile:=True, _
                DialogTitle:="Please select an input file...", _
                flags:=ahtOFN_HIDEREADONLY)

End Sub


Private Sub Form_Current()
'synchronize the Client or Deal form if it is open
 If ranCurrent = True Then Exit Sub     ' prevents double execution
    ranCurrent = True
    If IsOpen("ClientsView") Then
        clientRef.DefaultValue = Forms!ClientsView.clientID
    End If

    If IsOpen("Deals_view") Then
        dealRef.DefaultValue = Forms!deals_view.dealID
        clientRef.DefaultValue = Forms!deals_view.clientRef
    End If

End Sub
Private Sub AddFiletypeBtn_Click()
On Error GoTo Err_AddFiletypeBtn_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FileType_Add"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_AddFiletypeBtn_Click:
    Exit Sub

Err_AddFiletypeBtn_Click:
    MsgBox Err.Description
    Resume Exit_AddFiletypeBtn_Click

End Sub
