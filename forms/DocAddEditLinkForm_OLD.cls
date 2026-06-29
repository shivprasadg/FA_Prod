Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RecID As Integer
Dim Committ As String
Dim SQLz, LoadArg As String
Dim FormX, RequestX As String


Private Sub AddLink_Click()
    'DoCmd.GotoRecord , , acNewRec
    Me.ClientID = Split(LoadArg, "|")(0)
    Me.MLID = Split(LoadArg, "|")(1)
    Me.SchID = Split(LoadArg, "|")(2)
    Me.groupID = Split(LoadArg, "|")(3)
    Me.UnitID = Split(LoadArg, "|")(4)
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

Private Sub Form_Load()
    Me.Form.RecordSource = ""
    If Me.OpenArgs & "test" <> "test" Then
    LoadArg = Me.OpenArgs
    Call LoadData
    End If
End Sub
Public Sub LoadData()
        RequestX = Split(Split(LoadArg, "|")(6), ":")(0)
        FormX = Split(LoadArg, "|")(5)
        If RequestX = "Add" Then
            'DoCmd.GotoRecord , , acNewRec
            Me.ClientID = Split(LoadArg, "|")(0)
            Me.MLID = Split(LoadArg, "|")(1)
            Me.SchID = Split(LoadArg, "|")(2)
            Me.groupID = Split(LoadArg, "|")(3)
            Me.UnitID = Split(LoadArg, "|")(4)
            MsgBox "Copy and Paste 'SharePoint Link' to the: " & vbNewLine & "'Link' Field", vbInformation, "Copy & Paste"
        End If
        If RequestX = "Edit" Then
            RecID = Split(LoadArg, "Edit:")(1)
            SQLz = "Select * From DocLinks Where DocID=" & RecID
            Me.Form.RecordSource = SQLz
        End If
End Sub

Private Sub GetLinkInfo()
    Dim Linkx As String
    Dim CleanLinkx As String
    CleanLinkx = Me.dLink
    If CleanLinkx Like "*#*" Then CleanLinkx = SplitString(SplitString(CleanLinkx, "#", 1), "#", 0)

    Linkx = GetSharePointLink(CleanLinkx)
    If Linkx Like "*out of range *" Then
        MsgBox "This is not a normal SharePoint File", vbCritical, "Are you playing with me?"
        Exit Sub
    End If
    Me.dtags = SplitString(Linkx, "|", 1)
    Me.dName = SplitString(Linkx, "|", 2)
    Me.dtype = SplitString(Linkx, "|", 3)
    CleanLinkx = Me.dLink
    If CleanLinkx Like "*#*" Then CleanLinkx = SplitString(SplitString(CleanLinkx, "#", 1), "#", 0)
    Linkx = Me.dName & "#" & CleanLinkx & "#"
    Me.dLink = Linkx
    If RequestX = "add" Then Call InsertRecord
    Me.Form.Refresh
    Forms(FormX).DcLinksSub.Form.Requery
End Sub
Private Sub InsertRecord()
    SQLz = "INSERT INTO DocLinks (ClientID, MLID, SchID, GroupID, UnitID, dtags, dName, dtype, dlink) "
    SQLz = SQLz & "Select " & Me.ClientID & "," & Me.MLID & "," & Me.SchID & "," & Me.groupID & "," & Me.UnitID & ",'" & Me.dtags & "',"
    SQLz = SQLz & "'" & Me.dName & "', '" & Me.dtype & "', '" & Me.dLink & "';"
    DoCmd.SetWarnings False
    'DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

End Sub
Private Sub GetLink_Click()
    Call GetLinkInfo
End Sub

Private Sub SaveandClose_Click()
    If Me.dLink & "test" = "test" Then Exit Sub
       Forms(FormX).DcLinksSub.Form.Requery
       DoCmd.close , Form(Me.Form.Name)
    '    If IsOpen("GroupForm") Then Forms!(FormX).DcLinksSub.Form.Requery
    '    If IsOpen("Schedule_Form") Then Forms!Schedule_Form.DcLinksSub.Form.Requery
    '    If IsOpen("MstrLease") Then Forms!MstrLease.DcLinksSub.Form.Requery

End Sub
