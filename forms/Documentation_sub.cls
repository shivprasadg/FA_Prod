Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'------------------------------------------------------------
' AddContactBtn_Click
'
'------------------------------------------------------------
Private Sub AddContactBtn_Click()
On Error GoTo AddContactBtn_Click_Err

Dim stDocName As String
Dim stLinkCriteria As String

    stDocName = "Contacts_View"
    stLinkCriteria = "[clientRef]=" & Forms!deals_view![clientRef]

    DoCmd.OpenForm stDocName, , , stLinkCriteria

AddContactBtn_Click_Exit:
    Exit Sub

AddContactBtn_Click_Err:
    MsgBox Error$
    Resume AddContactBtn_Click_Exit

End Sub
