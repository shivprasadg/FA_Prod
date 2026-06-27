Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset
Dim BillingChanged As Boolean
Dim PreContactx As String
Dim msg1 As String
Dim SQLz       As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub AddContact_Click()
    DoCmd.GotoRecord , , acNewRec
    Me.clientID = Me.Parent.clientID
End Sub

Public Sub SetContacts()
    Me.ConObsoleteDt.Locked = True
    Me.ContactObsolete.Locked = True
    SQLz = "SELECT contactID, clientRef, contactClassification As Classification, contactFullName AS Contact FROM Location_contacts " & _
           "WHERE clientRef=" & Me.Parent.clientID & " Order by contactFirst, contactLast ;"
    Me.ContactsList.RowSource = SQLz
    Me.Form.RecordSource = "Select * From Location_Contacts Where clientRef=" & Me.Parent.clientID & ""
End Sub


Private Sub contactCell_AfterUpdate()
   If Me.contactClassification = "Billing" Then BillingChanged = True
   Call MakeEmailMsg
End Sub

 Private Sub contactClassification_AfterUpdate()
   If Me.contactClassification = "Billing" Then BillingChanged = True
   Call MakeEmailMsg
End Sub

Private Sub contactEmail_AfterUpdate()
   If Me.contactClassification = "Billing" Then BillingChanged = True
   Call MakeEmailMsg
End Sub

Private Sub contactFax_AfterUpdate()
   If Me.contactClassification = "Billing" Then BillingChanged = True
   Call MakeEmailMsg
End Sub

Private Sub contactPrimaryPhone_AfterUpdate()
   If Me.contactClassification = "Billing" Then BillingChanged = True
   Call MakeEmailMsg
End Sub

Private Sub Form_AfterUpdate()

End Sub

Public Sub MakeEmailMsg()
    If Me.contactClassification Like "*Billing*" And PreContactx <> "" And BillingChanged = True Then
        'make email message
        msg1 = ""
        msg1 = msg1 & "<html><body><p style='color:RED;font-family:Segoe UI;font-size:20px;'><b>Billing Contact Change</b></p>"
        msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Operations<br />"
        msg1 = msg1 & "Subject:&#160;&#160;&#160;&#160;&#160;&#160;Billing Contact Changed, Re Client:<b> " & Me.Parent.ClientShNm & "</b><br />"
        msg1 = msg1 & "Priority:<b>&#160;&#160;&#160;&#160;&#160;&#160;High </b><br/><br/>"
        msg1 = msg1 & "All,<br/><br/>There was a change submitted to a Billing Contact, the details are below.<br />"
        msg1 = msg1 & "Client:<b> " & Me.Parent.ClientShNm & "</b><br /><br />"
        msg1 = msg1 & "Previous Contact Details: <br />" & PreContactx & "<br /><br />"
        msg1 = msg1 & "New Contact Details:<br />"
        msg1 = msg1 & "" & Me.contactFullName & "</b><br />"
        msg1 = msg1 & "Ph:<b> " & Me.contactPrimaryPhone & "</b><br />"
        msg1 = msg1 & "Cell:<b> " & Me.contactCell & "</b><br />"
        msg1 = msg1 & "Fax:<b> " & Me.contactFax & "</b><br />"
        msg1 = msg1 & "Email:<b> " & Me.contactEmail & "</b><br />"
        msg1 = msg1 & "Type:<b> " & Me.contactClassification & "</b><br /><br />"
        msg1 = msg1 & "FA User who made change: <B>" & FAUserFullName & "</B><br /><br />"
        msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
        msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
        msg1 = msg1 & "Atlaas Core,<br /><br /><br />"
        msg1 = msg1 & "</body></html>"
     Else
        msg1 = ""
     End If

End Sub
Public Sub SendEmailx()
    If Me.contactClassification Like "*Billing*" And msg1 <> "" And PreContactx <> "" And BillingChanged = True Then
        BillingChanged = False
        Call SendEmailMessage("ljohnson@fleetadvantage.com", "Billing Contact Change", msg1, True, True, False, "ljohnson@fleetadvantage.com")
        PreContactx = ""
        msg1 = ""
    End If
End Sub
Public Sub Form_Current()
msg1 = ""
PreContactx = ""
On Error Resume Next
    If Me.contactClassification Like "*Billing*" Then
        PreContactx = Me.contactFullName
        PreContactx = PreContactx & vbNewLine & "Ph:<b> " & Me.contactPrimaryPhone & "</b>"
        PreContactx = PreContactx & vbNewLine & "Cell:<b> " & Me.contactCell & "</b>"
        PreContactx = PreContactx & vbNewLine & "Fax:<b> " & Me.contactFax & "</b>"
        PreContactx = PreContactx & vbNewLine & "Email:<b> " & Me.contactEmail & "</b>"
        PreContactx = PreContactx & vbNewLine & "Type:<b> " & Me.contactClassification & "</b>"
    End If
On Error GoTo 0
End Sub

Private Sub LocationCBO_Click()
    Dim Msgx As String
    Dim LinkTypeX As Integer
    Msgx = ""
    If IsNull(Me.LocationCbo.Column(0)) Then Msgx = Msgx & vbNewLine & "   Location"
    If IsNull(Me.contactID) Then Msgx = Msgx & vbNewLine & "   Contact to Link"
    If IsNull(Me.contactClassification) Then Msgx = Msgx & vbNewLine & "   Type of Link to use"
   LinkTypeX = DLookup("LinkTypeID", "ContactsLinkedToLocationTypes", "LinkType='" & Me.contactClassification & "'")
    If Msgx <> "" Then
        Msgx = "You must Select a" & Msgx
        MsgBox Msgx, vbCritical, "Missing Data"
        Exit Sub
    End If
    If DLookup("ContactID", "ContactsLinkedToLocation", "ContactID=" & Me.contactID & " AND LinkTypeID =" & LinkTypeX & " AND LocationID=" & Me.LocationCbo.Column(0)) Then
        MsgBox Me.contactFirst & " " & Me.contactLast & vbNewLine & "Is already linked  to " & Me.LocationCbo.Column(1) & vbNewLine & "as a '" & Me.contactClassification & "' contact type" & vbNewLine & vbNewLine & "Canceled...", vbCritical, "We've Done this already"
        Exit Sub
    End If

    SQLz = "INSERT INTO ContactsLinkedToLocation ( ContactID, LocationID, LinkTypeID ) " & _
           "SELECT " & Me.contactID & ", " & Me.LocationCbo.Column(0) & ", " & LinkTypeX & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LinkedLocationsList.Requery
    Me.Parent.Locations_sub.Form.LinkEdContactsList.Requery
    Me.ContactDetail = ""
End Sub
Private Sub UnLinkedContact_Click()
    If IsNull(Me.LinkedLocationsList.Column(0)) Then Exit Sub
    Dim Askit  As String
    Askit = MsgBox("You're about to Un-Link..." & vbNewLine & "Location: " & Me.LinkedLocationsList.Column(3) & vbNewLine & "From:  " & Me.contactFirst & " " & Me.contactLast & vbNewLine & "Type:   " & Me.LinkedLocationsList.Column(2) & vbNewLine & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Un-Link")
    If Askit <> 6 Then Exit Sub
    SQLz = "Delete From ContactsLinkedToLocation " & _
           "WHERE  LocationID=" & Me.LinkedLocationsList.Column(0) & " AND LinkTypeID=" & Me.LinkedLocationsList.Column(5) & " AND ContactID=" & Me.contactID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LinkedLocationsList.Requery
    Me.Parent.Locations_sub.Form.LinkEdContactsList.Requery
    Me.ContactDetail = ""

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
    Me.Form.RecordSource = "Select * From Location_Contacts Where ContactId=" & Me.ContactsList.Column(0) & ""
    Me.ContactDetail = ""
    SQLz = "SELECT LocationID, LocationName, LocationStreet, LAddrLine1, LAddrLine2, LCity, LState, LZip " & _
           "From Locations WHERE ClientGroupID=" & Me.Parent.ClientGroupID & " AND LocObsolete = False;"
    Me.LocationCbo.RowSource = SQLz
    SQLz = "SELECT LocationID, ContactID, LinkType, Locationsn & ' (' & [LocationType] & ')' AS Location, City & ', ' & State AS City-State, LinkTypeID "
    SQLz = SQLz & "FROM ContactsLinkedToLocationQuery "
    SQLz = SQLz & "WHERE C.ContactID =" & Me.contactID & " "
    SQLz = SQLz & "ORDER BY Locationsn AND LocationType;"
    Me.LinkedLocationsList.RowSource = SQLz
    Me.LinkedLocationsList = Me.ContactsList.Column(0)
    Call LinkedLocationsList_Click
    If Me.contactClassification = "Billing" And PreContactx <> "" Then
       Call MakeEmailMsg
       If msg1 <> "" Then Call SendEmailx
    End If
End Sub

Private Sub DeleteContact_Click()
    If IsNull(Me.contactFullName) Then Exit Sub
    Dim Response
    On Error GoTo FixMe
    Response = MsgBox("Are you sure you want to DELETE " & Me.contactFullName & "?", vbQuestion + vbYesNo + vbCritical, "DELETE!!")
    If Response = vbYes Then
        SQLz = "Delete * From Contacts WHERE ContactId = " & Me.contactID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        Me.ContactsList.Requery
        DoCmd.SetWarnings True
        Me.ContactsList = DMax("contactID", "Location_contacts", "clientRef=" & Me.Parent.clientID)
        Call ContactsList_Click
        Me.Form.Refresh
    Else
        Exit Sub
    End If
FixMe:
End Sub

Private Sub Form_Close()
   Call SendEmailx
   On Error GoTo Form_Close_Error
   On Error GoTo 0
   Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_ClientContact_sub"
End Sub



Private Sub LinkedLocationsList_Click()
On Error GoTo FixMe
    SQLz = "Select * From Locations WHERE LocationID =" & Me.LinkedLocationsList.Column(0) & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Me.ContactDetail = Me.[contactFullName] & vbNewLine & Me.contactTitle & vbNewLine & _
                       RST!LAddrLine1 & vbNewLine & _
                       RST![LCity] & " " & RST!LState & " " & RST!LZip & vbNewLine & _
                       "Phone: " & Left(Me.[contactPrimaryPhone], 3) & "-" & Mid(Me.[contactPrimaryPhone], 4, 3) & "-" & Right(Me.[contactPrimaryPhone], 4) & vbNewLine & _
                       "Email: " & Me.[contactEmail] & vbNewLine
    RST.close
FixMe:
    Set RST = Nothing
End Sub



Private Sub LocationCbo_GotFocus()
    If IsNull(Me.ContactsList.Column(0)) Or Me.Form.NewRecord Then
        Call SaveUpdate_Click
        Me.ContactsList = Me.contactID
        'Call Form_Current
    End If
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
    If IsNull(Me.contactFirst.Value) Then Msgx = Msgx & " First Name |"
    If IsNull(Me.contactLast.Value) Then Msgx = Msgx & " Last Name |"
    If Me.contactClassification.Value & "test" = "test" Then Msgx = Msgx & " Contact Class |"
    If Msgx <> "|" Then
        MsgBox "Please Enter Data for All Required Fields" & vbNewLine & "You're Missing: " & Msgx, , Reuired    '
        Exit Sub
    End If
    If Me.contactClassification = "Billing" And PreContactx <> "" And msg1 <> "" Then
       Call SendEmailx
       PreContactx = ""
    End If
       PreContactx = ""
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
    If Me.contactClassification = "Billing" Then BillingChanged = True
    Call MakeEmailMsg
End Sub
