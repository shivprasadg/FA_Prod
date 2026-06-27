Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub contactFirst_AfterUpdate()
    [contactFullName] = Trim([contactPrefix] & " " & [contactFirst] & " " & _
    [contactMiddle] & " " & [contactLast] & " " & [contactSuffix])

End Sub

Private Sub contactLast_AfterUpdate()
    [contactFullName] = Trim([contactPrefix] & " " & [contactFirst] & " " & _
    [contactMiddle] & " " & [contactLast] & " " & [contactSuffix])

End Sub

Private Sub contactMiddle_AfterUpdate()
    [contactFullName] = Trim([contactPrefix] & " " & [contactFirst] & " " & _
    [contactMiddle] & " " & [contactLast] & " " & [contactSuffix])

End Sub

Private Sub contactPrefix_AfterUpdate()
    [contactFullName] = Trim([contactPrefix] & " " & [contactFirst] & " " & _
    [contactMiddle] & " " & [contactLast] & " " & [contactSuffix])

End Sub

Private Sub contactSuffix_AfterUpdate()
    [contactFullName] = Trim([contactPrefix] & " " & [contactFirst] & " " & _
    [contactMiddle] & " " & [contactLast] & " " & [contactSuffix])

End Sub

Private Sub Form_Load()


If IsOpen("deals_view") Then
    clientRef = Forms!deals_view.clientRef
End If

End Sub
