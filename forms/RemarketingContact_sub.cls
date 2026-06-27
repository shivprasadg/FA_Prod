Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset
Dim SQLz       As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Public Sub AddContact_Click()
If Me.Parent.NewRecord Then
MsgBox "You must select a buyer first...", , "No Buyer Selected"
Exit Sub
End If
    DoCmd.GotoRecord , , acNewRec
    Me.clientID = Me.Parent.clientID
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



Public Sub ContactsList_Click()
SQLz = "Select * From Location_contacts where ContactID =" & Me.ContactsList.Column(0) & ";"
'Me.Form.Filter = ""
'Me.Form.FilterOn = False
'    On Error GoTo fixme
'    Set RST = Me.RecordsetClone
'    RST.FindFirst "ContactId = " & Me.ContactsList.Column(0) & ""
 '   Me.Bookmark = RST.Bookmark
    'Me.Form.Repaint
 Me.Form.RecordSource = SQLz
    Me.Parent.contactID = Me.ContactsList.Column(0)
   ' Me.Parent.EventHistory_Contact_Sub.Form.Repaint
  '  Me.Parent.EventHistory_Contact_Sub.Form.Refresh
    Exit Sub
FixMe:
End Sub

Private Sub DeleteContact_Click()
   If IsNull(Me.contactID) Then Exit Sub
    Dim Countx As Integer
    Countx = DCount("EventID", "Events", "ContactID=" & Nz(Me.contactID, 0))
    If Countx > 0 Then
        MsgBox Me.contactFirst & " " & Me.contactLast & " has " & Countx & " activities attached...," & vbNewLine & "This contact cannot be deleted.", vbCritical, "No Can Do"
        Exit Sub
    End If
    Dim Response
    On Error GoTo FixMe
    Response = MsgBox("Are you sure you want to DELETE " & Me.contactFirst & " " & Me.contactLast & "?", vbQuestion + vbYesNo + vbCritical, "DELETE!!")
    If Response = vbYes Then
        SQLz = "Delete * From Contacts WHERE ContactId = " & Me.contactID & ";"
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

Private Sub DoNotCall_Click()
If Me.DoNotCall = 0 And Me.ContactNote Like "*|contact is no longer with company|*" Then
Me.ContactNote = Replace(Me.ContactNote, "|contact is no longer with company|", "")
End If
End Sub
Private Sub contactOptOut_Click()
If Me.contactOptOut = 0 And Me.ContactNote Like "*|contact is no longer with company|*" Then
Me.ContactNote = Replace(Me.ContactNote, "|contact is no longer with company|", "")
End If
End Sub
Private Sub contactClassification_Click()
If Me.contactClassification <> "No Longer with Co." And Me.ContactNote Like "*|contact is no longer with company|*" Then
Me.ContactNote = Replace(Me.ContactNote, "|contact is no longer with company|", "")
End If
End Sub
Private Sub Form_Current()
    Me.ConObsoleteDt.Locked = True
    Me.ContactObsolete.Locked = True
    'SQLz = "SELECT contactID, clientRef, contactClassification, contactFullName AS Contact FROM Location_contacts " & _
     "WHERE clientRef=" & Me.Parent.clientID & " Order by contactFullName ;"
    Me.ContactsList.Requery
    'Me.Form.Repaint
End Sub

Private Sub MakeFullContactBtn_Click()
    Call MakeFullContact
End Sub

Private Sub Printcontacts_Click()
    DoCmd.OpenReport "CONTACTPRNT", acViewPreview
End Sub

Private Sub SaveUpdate_Click()
    Dim Msgx   As String
    Msgx = "|"
    If IsNull(Me.contactFirst.Value) Or Me.contactFirst Like "*Here*" Then Msgx = Msgx & " First Name |"
    If IsNull(Me.contactLast.Value) Or Me.contactLast Like "*Here*" Then Msgx = Msgx & " Last Name |"
    If Me.contactClassification.Value & "test" = "test" Then Msgx = Msgx & " Contact Class |"
    If Msgx <> "|" Then
        MsgBox "Please Enter Proper Data for All Required Fields" & vbNewLine & "You're Missing: " & Msgx, , Reuired    '
        Exit Sub
    End If
    Me.Form.Refresh
End Sub

Private Sub Unlockit_Click()
    Me.ConObsoleteDt.Locked = False
    Me.ContactObsolete.Locked = False
End Sub

Public Sub MakeFullContact()
    Me.contactFullName = Trim([contactPrefix] & " " & [contactFirst] & " " & _
                              Nz([contactMiddle], "") & " " & [contactLast] & " " & [contactSuffix])
    Me.contactFullName = Trim(Me.contactFullName)
End Sub
