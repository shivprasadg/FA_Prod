Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Dim OldNum, NewNum, NextNum As Integer
Dim RST As DAO.Recordset

Private Sub AddDrawDown_Click()
    DoCmd.GotoRecord , , acNewRec
    Me.CoverBox.Visible = False
    Me.DrawBankId = BankList.Column(0)
End Sub

Private Sub Clear_Search_Click()
Me.SearchClient = ""
Call BankClientTabs_Click
End Sub

Private Sub CloneADraw_Click()
Dim Response As String
NextNum = Nz(DLookup("DrawNumber", "DrawMaxNumber", "DrawRequestNumber='" & Me.DrawRequestNumber & "'"), 0) + 1
    Response = MsgBox("You're about to Clone" & _
               vbNewLine & "Draw #" & Me.DrawNumber & " Request: " & Me.DrawRequestNumber & " Bank: " & DrawBankId.Column(1) & _
               vbNewLine & "As..." & _
               vbNewLine & "Draw #" & NextNum & " Request: " & Me.DrawRequestNumber & " Bank: " & DrawBankId.Column(1) & _
               vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "Is Cloning Even Legal?")
       If Response = vbNo Then Exit Sub
SQLz = "INSERT INTO DrawDowns ( DrawBankId, MLID, ClientID, DrawTypeId, DrawBankAccountNumber, DrawCreditLimit, DrawRequestNumber, CreatedBy, CreatedOn, DrawName, DrawNumber ) " & _
       "SELECT " & Me.DrawBankId & " as Bank, " & Me.MLIDCbo & " AS MLease, " & Me.clientID & " AS Client, " & Me.DrawTypeId & " AS type, '" & Me.DrawBankAccountNumber & "' AS Account, '" & Me.DrawCreditLimit & "' AS Limit, " & _
       "'" & Me.DrawRequestNumber & "' AS Request, '" & Forms!FAMainMenu.USERID & "' AS User, #" & Date & "# AS Created, '" & Me.DrawName & "' AS Name, " & NextNum & " AS Dnumber;"

DoCmd.RunSQL (SQLz)
Me.DrawDownList.Requery
Me.Form.Requery
Me.Form.Repaint
MsgBox "Draw Cloned", , "Draw inserted"
End Sub



Private Sub DrawTypeId_Click()
If Me.DrawTypeId = 6 Then
  SQLz = "SELECT clientID, ClientShortName AS Client FROM vw_SixKeys WHERE MLOrig<>'FA' GROUP BY clientID, ClientShortName ORDER BY ClientShortName;"
Else
  SQLz = "SELECT clientID, ClientShortName AS Client FROM vw_SixKeys WHERE MLOrig='FA' GROUP BY clientID, ClientShortName ORDER BY ClientShortName;"
End If
Me.ClientCbo.RowSource = SQLz
End Sub

Public Sub InvoiceList_Click()
    SQLz = "SELECT * FROM DrawDowns_Active Where DrawDownId=" & Me.InvoiceList & " Order BY MLID, DrawRequestNumber, DrawNumber DESC;"
    Me.DrawDownList.RowSource = SQLz
    Me.CoverBox.Visible = True
End Sub
Private Sub ClientCBO_Click()
If Me.DrawTypeId = 6 Then
    SQLz = "Select MLID, MLNO From MstrLease Where MLOrig <> 'FA' and Client=" & Me.ClientCbo & " Order BY MLID " 'DrawRequestNumber, DrawNumber DESC;"
Else
    SQLz = "Select MLID, MLNO From MstrLease Where MLOrig = 'FA' and Client=" & Me.ClientCbo & " Order BY MLID " ', DrawRequestNumber, DrawNumber DESC;"
End If
    Me.MLIDCbo.RowSource = SQLz
End Sub
Private Sub ClientsList_Click()
    SQLz = "SELECT * FROM DrawDowns_Active Where ClientID=" & Me.ClientsList & " Order BY DateDrawn Desc;"
    Me.DrawDownList.RowSource = SQLz
    Me.Form.RecordSource = "Select * From DrawDowns Where ClientID=" & Me.ClientsList.Column(0)
    Me.CoverBox.Visible = True
End Sub
Public Sub BankList_Click()
    SQLz = "Select * From DrawDowns_Active Where BankId=" & Me.BankList & " Order DrawNumber Desc, DrawRequestNumber;"
    Me.Form.RecordSource = "Select * From DrawDowns Where DrawBankID=" & Me.BankList.Column(0) & ";"
    Me.DrawDownList.RowSource = SQLz
    Me.CoverBox.Visible = True
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Public Sub DrawDownList_Click()
'On Error GoTo Fixme
'    Set RST = Me.RecordsetClone
'    RST.FindFirst "DrawDownId = " & Me.DrawDownList.Column(1) & ""
'    Me.Bookmark = RST.Bookmark
'    RST.close
'    Set RST = Nothing
'Me.CoverBox.Visible = False
Me.Form.RecordSource = "Select * From DrawDowns Where DrawdownID=" & Me.DrawDownList.Column(1)
Me.ClientCbo.Requery
Me.MLIDCbo.Requery
FixMe:
End Sub
Public Sub DrawNumber_GotFocus()
OldNum = Nz(Me.DrawNumber, 0)
End Sub
Public Sub DrawNumber_Change()
    On Error GoTo Grabnull
    If Not IsNumeric(Me.DrawNumber.Text) Then
    MsgBox "Sorry, You MUST insert a 'vaild number' for this drawdown.", vbInformation, "Numbers Only"
    If OldNum > 0 Then Me.DrawNumber = OldNum Else Me.DrawNumber = Null
    Exit Sub
    End If
    'Call DrawNumber_BeforeUpdate
Grabnull:

End Sub
'---------------------------------------------------------------------------------------
' Procedure : DrawNumber_BeforeUpdate
' Author    : ljohnson
' Date      : 11/10/2014
' Purpose   : make sure there are no dup in DrawDown numbers
'---------------------------------------------------------------------------------------
'
Public Sub DrawNumber_AfterUpdate()
If Me.DrawRequestNumber & "test" = "test" Then
MsgBox "Sorry, but you must attach this Draw to:" & vbNewLine & "   A Leasee" & vbNewLine & "  A Master Lease " & vbNewLine & "Before you can assign a Draw Number", vbInformation, "Need more info"
Exit Sub
End If

   If DCount("DrawRequestNumber", "Drawdowns", "DrawRequestNumber = '" & Me.DrawRequestNumber & "' AND DrawNumber = " & Me.DrawNumber) > 0 Then
        MsgBox "There aready is a Draw Number '" & Me.DrawNumber & "' for Request: " & Me.DrawRequestNumber, , "No Dups"
        If OldNum > 0 Then Me.DrawNumber = OldNum Else Me.DrawNumber = Null
    Exit Sub
    End If
Me.DrawDownList.Requery
nodup:
    Call MarkUpdate
End Sub

Private Sub EditMe_Click()
    Call Me.Unlockit
End Sub
Public Sub Unlockit()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "EditMe" Then
            Me(Ccrl.Name).Locked = False
        End If
    Next Ccrl
    Me.EditLabel.caption = "Draw Details (Un-locked)"
End Sub
Public Sub LockIt()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "EditMe" Then
            Me(Ccrl.Name).Locked = True
        End If
    Next Ccrl
    Me.EditLabel.caption = "Draw Details (Locked)"
End Sub

Private Sub OpenDrawDownManager_Click()
DoCmd.OpenForm "DrawDownManager"
End Sub

Private Sub SaveUpdate_Click()
 If IsNull(Me.CreatedOn) Then
        Me.CreatedOn = Date
        Me.CreatedBy = SplitString(Forms!FAMainMenu.USERID, " ", 0)
  End If
    'me.MLIDCbo = Me.FAID.Column(1)
    If IsNull(Me.DrawNumber) Then
       NextNum = Nz(DLookup("DrawNumber", "DrawMaxNumber", "DrawRequestNumber='" & Me.DrawRequestNumber & "'"), 0)
    Me.DrawNumber = Nz(NextNum, 0) + 1
    End If
    'Call MarkUpdate

    Me.Form.Refresh
    Me.DrawDownList.Requery
    Call LockIt
End Sub
Private Sub MarkUpdate()
    Me.ModifiedBy = Forms!FAMainMenu.USERID
    Me.ModifiedOn = Now()
End Sub


Private Sub Form_Current()
    If Me.Form.NewRecord Then
    Call Unlockit
    Me.CloneADraw.Enabled = False
    Else
    Me.CloneADraw.Enabled = True
    Call LockIt
    End If
    If Me.DrawTypeId = 6 Then Me.RequestLabel1.caption = "Loan Num" Else Me.RequestLabel1.caption = "Request Num"
    If Me.DrawTypeId = 6 Then Me.RequestLabel2.caption = "Loan Number" Else Me.RequestLabel2.caption = "Request Number"
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub
Private Sub BankClientTabs_Click()
    'If Me.SearchClient <> "" Then
        If Me.BankClientTabs = 1 Then
            SQLz = "Select * From DrawClientsList Where Client Like '*" & Me.SearchClient & "*' Order By Client;"
            Me.ClientsList.RowSource = SQLz
        End If
        If Me.BankClientTabs = 0 Then
            SQLz = "SELECT * From DrawBanksList  Where Bank Like '*" & Me.SearchClient & "*' Order By Bank;"
            Me.BankList.RowSource = SQLz
        End If
        If Me.BankClientTabs = 2 Then
            SQLz = "SELECT * From DrawDowns_Invoices_Drawn Where Invoice Like '*" & Me.SearchClient & "*' Or Client Like '*" & Me.SearchClient & "*' Order By Invoice;"
            Me.InvoiceList.RowSource = SQLz
        End If
    'End If
End Sub
Public Sub ByClient_Click()
Me.SearchClient = ""
Call SearchClient_Change
End Sub
Private Sub ByBank_Click()
Me.SearchClient = ""
Call SearchClient_Change
End Sub

Private Sub SearchClient_Change()
    If Me.BankClientTabs = 1 Then
        SQLz = "Select * From DrawClientsList Where Client Like '*" & Me.SearchClient.Text & "*' Order By Client;"
        Me.ClientsList.RowSource = SQLz
    End If
    If Me.BankClientTabs = 0 Then
        SQLz = "SELECT * From DrawBanksList  Where Bank Like '*" & Me.SearchClient.Text & "*' Order By Bank;"
        Me.BankList.RowSource = SQLz
    End If
    If Me.BankClientTabs = 2 Then
        SQLz = "SELECT * From DrawDowns_Invoices_Drawn Where Invoice Like '*" & Me.SearchClient.Text & "*' Or Client Like '*" & Me.SearchClient & "*' Order By Invoice;"
        Me.InvoiceList.RowSource = SQLz
    End If
End Sub
