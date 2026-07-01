Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, UPDwn As String
Dim x As Integer, Y As Integer, cx As Integer, cy As Integer
Public Sub OpenTarget(TargetX As String)
'   DoCmd.Minimize
    DoCmd.OpenForm TargetX, acNormal, , , acFormEdit, , Me.Form.Name & ":BC"
End Sub

Private Sub ATLAASlogo_DblClick(Cancel As Integer)
    If FAUserAccessLevel = 99 Then DoCmd.OpenForm "ITFunctions"
End Sub

Private Sub DocsMenu_Click()
    DoCmd.close acForm, "DocsMenu'"
    DoCmd.OpenForm "DocsMenu"
End Sub

Private Sub Form_Load()
    Call CreateShortCutMenus
    On Error GoTo MoveOn
    If DevOrPro <> "" Then Me.DevLabel.caption = DevOrPro
    Dim PiX, TwipX As Integer    ' X = HORIZONTAL
    Dim PiY, TwipY As Integer    ' Y = VERTICAL
    ''get Screensize from functions
    TwipX = fPixelsToTwips(GetSystemMetrics(0), DIRECTION_HORIZONTAL)
    TwipY = fPixelsToTwips(GetSystemMetrics(1), DIRECTION_VERTICAL)
    DoCmd.MoveSize (TwipX - (Me.Form.Width + 200)), 0
    Call MinMax_Click

MoveOn:
    If DLookup("[Active User]", "TBL_Employees", "[User ID] = '" & LCase(Environ("username")) & "'") = True Then
        GoTo AccessGranted
    Else
        MsgBox "You are not authorized to use the Fleet Advantage Data Base at this time. Please contact technical support to gain access", vbCritical, "LOGIN"
        DoCmd.CloseDatabase
    End If
AccessGranted:
    Me.UserName = FAUserFullName
    Me.USERID = FaUserName & " (" & FAUserAccessLevel & ")"
    Me.Version = AtlaasVersion
    Me.Version.ControlTipText = CurrentDb.Name
    'Call CustomMenus ' no need to have a start up menu
    Call Form_Timer
    Call CheckSpecialUser
FixMe:
End Sub
Private Sub CheckSpecialUser()
If FAUserAccessLevel = 99 Or FAUserIDx = 27 Then
Me.ScheduleControl.Width = 1920
Me.TreeCommand.Visible = True
Else
Me.ScheduleControl.Width = 2975
Me.TreeCommand.Visible = False
End If


End Sub
Private Sub Form_Timer()
If CInt(Left(Time(), 1)) >= 3 And Right(Time(), 2) = "PM" And CheckIfSendBLCDEmailSent = False Then Call BLCDSyndFinalMessage(0)
End Sub

Private Sub MinMax_Click()
    If UPDwn = "Down" Then
        x = 0
        Y = 0
        cx = 1800
        cy = 1550
        xg_SizeWindow "Access", x, Y, cx, cy
        CommandBars.ActiveMenuBar.Enabled = True

        UPDwn = "Up"
    Else
        x = 1700
        Y = 0
        cx = 50
        cy = 50
        xg_SizeWindow "Access", x, Y, cx, cy
        Me.UserName.SetFocus
        UPDwn = "Down"
    End If
End Sub

Private Sub ScheduleControl_Click()
    DoCmd.close acForm, "ScheduleControl"
    DoCmd.OpenForm "ScheduleControl"
End Sub

Private Sub UpdateMe_Click()
    Dim BatchNameX As String
    BatchNameX = "C:\FAAS\UpdateMe.bat"
    Shell BatchNameX
    DoCmd.Quit
End Sub
Private Sub Clients_Prospects_Click()
DoCmd.OpenForm "ClientsView", , , , , , 4400
'  Call OpenTarget("ClientsView")
End Sub

Private Sub TreeCommand_Click()
DoCmd.OpenForm "TreeCommand"
End Sub

Private Sub View_Transactions_Click()
    Call OpenTarget("Deals_View")
End Sub

Private Sub Search_Transactions_Click()
    Call OpenTarget("Search_Deals")
End Sub

Private Sub LORPTM_Click()
    Call OpenTarget("LORPTM")
End Sub
Private Sub Accounting_Click()
 If FAUserAccessGroup Like "*ACCT*" Or FAUserAccessGroup Like "*ALL*" Then
  Call OpenTarget("AcctMenu")
Else
 MsgBox "You Are Not Authorized", vbCritical, "Access Denied"
End If
End Sub
Private Sub Management_Click()
    Call OpenTarget("ManagementMenu")
End Sub
Private Sub Operations_Click()
    Call OpenTarget("OPS Main Menu")
End Sub

Private Sub Search_Clients_Click()
    Call OpenTarget("Search_Clients")
End Sub

Private Sub ReMarket_Click()
    Call OpenTarget("Rmkt Menu")
End Sub
Private Sub UnitLookUp_Click()
    Call OpenTarget("UnitReverseLookUpForm")
End Sub

Private Sub ExitMe_Click()
'DoCmd.CloseDatabase
    Application.SetOption "Auto compact", False
    Dim Response, FormNameX As String
    Dim Cancel As Boolean
    FormNameX = Me.Form.Name
    Response = MsgBox("Are you sure you want to EXIT?", vbQuestion + vbYesNo + vbCritical, "Close ATLAAS Core?")
    If Response = vbYes Then
        Application.SetOption "Auto compact", False
        If Format(Date, "DDD") Like "Thu" Then Application.SetOption "Auto compact", True
        DoCmd.Quit
    End If
End Sub

Private Sub CustomMenus()
    Call LogUsersin
    Exit Sub
    If FAUserAccessLevel = 2 Then DoCmd.OpenForm "LO Menu"
    If FAUserAccessLevel = 3 Then DoCmd.OpenForm "AcctMenu"
    If FAUserAccessLevel = 4 Then DoCmd.OpenForm "AM_Menu"
    If FAUserAccessLevel = 5 Then DoCmd.OpenForm "DocsMenu"
    If FAUserAccessLevel = 6 Then DoCmd.OpenForm "ManagementMenu"
    If FAUserAccessLevel = 7 Then DoCmd.OpenForm "OPS VP Menu"
    If FAUserAccessLevel = 8 Then DoCmd.OpenForm "Marketing VP Menu"
    If FAUserAccessLevel = 10 Then DoCmd.OpenForm "Rmkt Menu"
    If FAUserAccessLevel = 11 Then DoCmd.OpenForm "ManagementMenuPF"
End Sub
