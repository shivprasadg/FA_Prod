Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz As String
Public RST As DAO.Recordset

Private Sub BuyerName_DblClick(Cancel As Integer)
    Call Unlockit
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    Exit Sub
CloseMe:
    On Error Resume Next
    DoCmd.close acForm, Me.Form.Name
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_BuyersForm"
    GoTo CloseMe
End Sub

Private Sub AddNewBuyer_Click()
    Dim Response, SQLset As String
    Response = MsgBox("You're about to ADD a NEW Buyer, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Bring me more Buyers")
    If Response = vbNo Then Exit Sub
    DoCmd.GotoRecord , , acNewRec
    Me.contactID = Null
    Me.ContactDetails.Form.RecordSource = "Select * From Location_contacts where ContactID =0;"
    Me.ContactDetails.Form.ContactsList.RowSource = ""
    Call Unlockit
End Sub

Public Sub BuyerList_Click()
    SQLz = "Select * From Buyer WHERE BuyerID = " & Nz(Me.BuyerList, Me.BuyerID) & ";"
    Me.Form.RecordSource = SQLz

    SQLz = "SELECT contactID, clientRef, contactClassification, contactFullName AS Contact, [clientRef] " & _
           "FROM Location_contacts WHERE clientRef=" & Nz(Me.BuyerList, Me.BuyerID) & " ORDER BY contactFullName;"
    Me.ContactDetails.Form.ContactsList.RowSource = SQLz
    Me.RemarketingInvoiceRequest_Sub.Form.BuyersOrderUnits_Sub.Form.Visible = False
FixMe:
    On Error GoTo 0
End Sub

Private Sub BuyerShortName_DblClick(Cancel As Integer)
    Call Unlockit
End Sub

Private Sub Unlockit()
    Me.BuyerName.Locked = False
    Me.BuyerShortName.Locked = False
End Sub

Private Sub LockIt()
    Me.BuyerName.Locked = True
    Me.BuyerShortName.Locked = True
End Sub

Private Sub CloseformBtn_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub SaveMe_Click()
    On Error GoTo SaveMe_Click_Error
    If Me.Dirty Then
        Me.Dirty = False
        MsgBox "Buyer '" & Me.BuyerName & "'" & vbNewLine & "was saved", vbInformation, AtlaasVersion
    End If
    On Error GoTo -1
    Exit Sub
SaveMe_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveMe_Click of VBA Document Form_BuyersForm"
End Sub

Private Sub SearchBuyer_Change()
    SQLz = "SELECT BuyerID, BuyerShortName AS Buyer FROM Buyer " & _
           "WHERE BuyerShortName Like '*" & Me.SearchBuyer.Text & "*'" & _
           "Or BuyerName Like '*" & Me.SearchBuyer.Text & "*' ORDER BY BuyerShortName;"
    '"WHERE clientCompanyName Like '*" & Me.SearchBuyer.Text & "*' ORDER BY clientCompanyName;"
    Me.BuyerList.RowSource = SQLz
End Sub
