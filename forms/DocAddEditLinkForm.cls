Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RecID As Integer
Dim Committ As String
Dim SQLz As String
Dim FormX As String
Dim RequestX As String


Private Sub AddLink_Click()
' DoCmd.GotoRecord , , acNewRec
    Me.clientID = Split(Me.OpenArgs, "|")(0)
    Me.MLID = Split(Me.OpenArgs, "|")(1)
    Me.SchID = Split(Me.OpenArgs, "|")(2)
    Me.groupID = Split(Me.OpenArgs, "|")(3)
    Me.UnitID = Split(Me.OpenArgs, "|")(4)
    Me.dLink = ""
    Me.dtags = ""
    Me.dName = ""
    Me.dtype = ""
    MsgBox "Copy and Past 'SharePoint Link' to the: " & vbNewLine & "'Link' Field", vbInformation, "How are you?"
End Sub

Private Sub DeleteLink_Click()
    Committ = MsgBox("Are you sure you want DELETE:" & vbNewLine & SplitString(Me.dLink, "#", 0) & "?", vbQuestion + vbYesNo, "Don't play with matches")
    If Committ <> "6" Then Exit Sub
    SQLz = "Delete From DocLinks Where DocID=" & Me.DocID
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
End Sub

Private Sub dLink_AfterUpdate()

    If Me.dLink & "test" = "test" Then Exit Sub
    If Me.dName & "test" <> "test" Then
        Committ = MsgBox("Are you sure you want change this Linked Doc?", vbQuestion + vbYesNo, "Please confrim")
    End If
    If Me.dName & "test" = "test" Or Committ = "6" Then Call GetLinkInfo
End Sub

Public Sub Form_Load()
    If Me.OpenArgs & "test" <> "test" Then
        RequestX = Split(Me.OpenArgs, "|")(6)
        FormX = Split(Me.OpenArgs, "|")(5)
        If RequestX = "Add" Then
            DoCmd.GotoRecord , , acNewRec
            Me.clientID = Split(Me.OpenArgs, "|")(0)
            Me.MLID = Split(Me.OpenArgs, "|")(1)
            Me.SchID = Split(Me.OpenArgs, "|")(2)
            Me.groupID = Split(Me.OpenArgs, "|")(3)
            Me.UnitID = Split(Me.OpenArgs, "|")(4)
            Me.ScheduleDesc = "Sch " & DLookup("Schno", "Schedule", "SchId=" & Me.SchID) & " + " & DLookup("SchShortDesc", "Schedule", "SchId=" & Me.SchID)
           End If
        If RequestX = "Edit" Then
            'RecID = Split(Me.OpenArgs, "Edit:")(1)
            'SQLz = "Select * From DocLinks Where DocID=" & RecID & ";"
            'Me.Form.RecordSource = SQLz
            'Me.Form.Requery
        End If
    End If
End Sub

Private Sub GetLinkInfo(Optional xDir As Boolean)
    Dim Linkx As String
    Dim CleanLinkx As String
    CleanLinkx = Me.dLink
    If InStr(1, CleanLinkx, "#") > 0 Then
        CleanLinkx = SplitString(SplitString(CleanLinkx, "#", 1), "#", 0)
    End If

    Linkx = GetSharePointLink(CleanLinkx, xDir)
    If Linkx Like "*out of range *" Then
        MsgBox "This is not a normal SharePoint File", vbCritical, "Are you playing with me?"
        Exit Sub
    End If

    Me.dtags = SplitString(Linkx, "|", 1)
    Me.dName = SplitString(Linkx, "|", 2)
    Me.dtype = SplitString(Linkx, "|", 3)
    CleanLinkx = Me.dLink
    Dim i As Integer
    Dim HasHash As Boolean
    For i = 1 To Len(CleanLinkx)
    If Mid(i, 1) = Chr(35) Then HasHash = True Else HasHash = False
    Next
    If HasHash = True Then
       CleanLinkx = SplitString(SplitString(CleanLinkx, "#", 1), "#", 0)
    End If
  '  Linkx = Me.dName & "#" & CleanLinkx & "#"
    If InStr(1, CleanLinkx, "#") > 0 Then CleanLinkx = SplitString(SplitString(CleanLinkx, "#", 1), "#", 0)
    If InStr(1, Me.dName, "#") > 0 Then Me.dName = SplitString(SplitString(Me.dName, "#", 1), "#", 0)

    Linkx = Me.dName & "#" & CleanLinkx & "#"
    Me.dLink = Linkx
    If RequestX = "add" Then Call InsertRecord
     Me.Form.Refresh
    Call SaveandClose_Click
End Sub
Private Sub InsertRecord()
    SQLz = "INSERT INTO DocLinks (ClientID, MLID, SchID, GroupID, UnitID, dtags, dName, dtype, dlink) "
    SQLz = SQLz & "Select " & Me.clientID & "," & Me.MLID & "," & Me.SchID & "," & Me.groupID & "," & Me.UnitID & ",'" & Me.dtags & "',"
    SQLz = SQLz & "'" & Me.dName & "', '" & Me.dtype & "', '" & Me.dLink & "';"
    DoCmd.SetWarnings False
    'DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

End Sub
Private Sub GetLink_Click()
    Call GetLinkInfo
    Call DoRefresh
End Sub

Private Sub SaveandClose_Click()
    If Me.Form.Dirty = True Then Me.Form.Dirty = False
    If Me.dLink & "test" = "test" Or Me.dName & "test" = "test" Then
        Committ = MsgBox("There isn't a Link" & vbNewLine & "Do you want to abort?", vbQuestion + vbYesNo, "Please confrim")
        If Committ = "6" Then
                SQLz = "Delete From DocLinks Where isnull(dLink)"
                DoCmd.SetWarnings False
                DoCmd.RunSQL (SQLz)
                DoCmd.SetWarnings True
            End If
        End If
    If FormX <> "" Then Forms(FormX).DocLinksSub.Form.Requery
    Call DoRefresh
        DoCmd.close acForm, Me.Form.Name
End Sub
Private Sub DoRefresh()
    If IsOpen("GroupForm") Then Forms!GroupForm.DocLinksSub.Form.Requery
    If IsOpen("Schedule_Form") Then Forms!Schedule_Form.DocLinksSub.Form.Requery
    If IsOpen("MstrLease") Then Forms!MstrLease.DocLinksSub.Form.Requery
    If IsOpen("LinkedDocsManager") Then Forms!LinkedDocsManager.Form.ScheduleList.Requery

End Sub

Private Sub TryDirectory_Click()
Me.dLink = GetSchedulDirectory(Me.ScheduleDesc)
If Me.dLink <> "" Then Call GetLinkInfo(True)
End Sub
