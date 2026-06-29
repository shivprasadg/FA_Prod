Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST As DAO.Recordset
Dim SQLz As String

Private Sub CloseMe_Click()
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub AddContact_Click()
 DoCmd.GotoRecord , , acNewRec
 Me.ContactsList.Requery
 Me.ClientID = Me.Parent.VendorID
End Sub

Private Sub contactFirst_AfterUpdate()
Call MakeFullContact
End Sub

Private Sub contactLast_AfterUpdate()
Call MakeFullContact
End Sub

Private Sub contactMiddle_AfterUpdate()
Call MakeFullContact
End Sub
Private Sub contactPrefix_AfterUpdate()
Call MakeFullContact
End Sub
Private Sub ContactObsolete_Click()
Me.ConObsoleteDt = Date
End Sub



Private Sub ContactsList_Click()
  'Me.Form.Filter = ""
  'Me.Form.FilterOn = False
  Set RST = Me.RecordsetClone
  RST.FindFirst "ContactId = " & Me.ContactsList.Column(0) & ""
  Me.Bookmark = RST.Bookmark
  Me.Form.Repaint
End Sub

Private Sub DeleteContact_Click()
If IsNull(Me.contactFullName) Then Exit Sub
  Dim Response
  On Error GoTo FixMe
      Response = MsgBox("Are you sure you want to DELETE " & Me.contactFullName & "?", vbQuestion + vbYesNo + vbCritical, "DELETE!!")
         If Response = vbYes Then
            SQLz = "Delete * From VendorContacts WHERE ContactId = " & Me.contactID & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            Me.ContactsList.Requery
            DoCmd.SetWarnings True
            Me.Form.Refresh
         Else
         Exit Sub
         End If
FixMe:
End Sub

Private Sub Form_Current()
Me.ConObsoleteDt.Locked = True
Me.ContactObsolete.Locked = True
SQLz = "SELECT contactID, VendorRef, ContactTitle as Title, [contactFirst] & [ContactLast] AS Contact " & _
       "FROM VendorContacts " & _
       "WHERE VendorRef=" & Me.Parent.VendorID & " ORDER BY ContactLast;"
Me.ContactsList.RowSource = SQLz
MakeFullContact
End Sub

Private Sub MakeFullContactBtn_Click()
Call MakeFullContact
End Sub

Private Sub Printcontacts_Click()
DoCmd.OpenReport "CONTACTPRNT", acViewPreview
End Sub

Private Sub SaveUpdate_Click()
Dim Msgx As String
Msgx = "|"
    If IsNull(Me.contactFirst.Value) Then Msgx = Msgx & " First Name |"
    If IsNull(Me.contactLast.Value) Then Msgx = Msgx & " Last Name |"
  ' If Me.contactClassification.Value & "test" = "test" Then msgx = msgx & " Contact Class |"
    If Msgx <> "|" Then
        MsgBox "Please Enter Data for All Required Fields" & vbNewLine & "You're Missing: " & Msgx, , Reuired '
        Exit Sub
    End If
Me.Form.Refresh

End Sub

Private Sub Unlockit_Click()
Me.ConObsoleteDt.Locked = False
Me.ContactObsolete.Locked = False
End Sub

Public Sub MakeFullContact()
    Me.contactFullName = [contactFirst] & " " & _
    Nz([contactMiddle], "") & " " & Trim([contactLast])
    Me.contactFullName = Trim(Me.contactFullName)
End Sub
