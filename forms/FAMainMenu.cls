Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, UPDwn, HidShw As String
Dim x As Integer, Y As Integer, cx As Integer, cy As Integer
Public Sub OpenTarget(TargetX As String)
'   DoCmd.Minimize
   If (TargetX = "SignOffDashboard" _
       Or TargetX = "OPS Main Menu" _
       Or TargetX = "LORPTM" _
       Or TargetX = "AcctMenu" _
       Or TargetX = "ManagementMenu") _
       And (FAUserAccessLevel < 11 Or FAUserAccessGroup Like "*BDEX*") Then
   MsgBox "You are not authorized to use " & TargetX & ". Please contact technical support to gain access", vbCritical, "Stop"
Else
    DoCmd.OpenForm TargetX, acNormal, , , acFormEdit, , Me.Form.Name & ":BC"
End If
End Sub

Private Sub ATLAASlogo_DblClick(Cancel As Integer)
    If FAUserAccess = 99 Then DoCmd.OpenForm "ITFunctions"
End Sub

Private Sub DocsMenu_Click()
    DoCmd.close acForm, "DocsMenu'"
    DoCmd.OpenForm "DocsMenu"
End Sub

Private Sub EventList_DblClick(Cancel As Integer)
'    Select Case Me.EventList.Column(6)
'    Case "Sch"
'        DoCmd.OpenForm "Schedule_Form", , , "SchID=" & Me.EventList.Column(2)
'    Case "Grp"
'        DoCmd.OpenForm "GroupForm", , , "SGrpID=" & Me.EventList.Column(3)
'    Case "RMKT"
'        Dim SQLz As String
'        Dim Cntx As Integer
'       On Error Resume Next
'        Cntx = CInt(Trim(SplitString((Trim(SplitString(Me.EventList.Column(0), "(", 1))), ")", 0)))
'        On Error GoTo 0
'        SQLz = "SELECT SourceCompany, [Sch|Grp], UnitID, UnitNum, UnitVin, Cstr(UnitYr) + ' ' + UnitMake + ' ' + UnitModel + ' ' + UnitSubType AS UnitDesc, " & _
'               "Buyer, InvoicePaid, DaysInInventory, UnitSalePrice " & _
'               "From RemarketingReportAllUnits WHERE InvoicePaid =#" & CDate(Me.EventList.Column(5)) & "# ORDER BY SourceCompany, [Sch|Grp], UnitID"
'        Call MessageUser("Open", "Open Quick Sold Report", "...just a second or two...")
'        Call SendReport2Excel(SQLz, Cntx & " Sold", Cntx & " Units that Sold on " & Me.EventList.Column(5), True, "|I,CB|J,$E|", Cntx, True)
'        Call MessageUser("Close")
'    End Select
End Sub

Private Sub Form_Load()
    Dim WhereX As String
    Dim PiX, TwipX As Integer    ' X = HORIZONTAL
    Dim PiY, TwipY As Integer    ' Y = VERTICAL
    Dim ScnX As Long
    DoCmd.SetWarnings True
    WhereX = WhereAmI
    If WhereX Like "*Phoenix*" Then
        Call ChangeConnectionString("FA-Phoenix", , "TBL_Employees")
    Else
        Call ChangeConnectionString(, , "TBL_Employees")
    End If

    'Call WhereAmI(1)
    'Call ChangeConnectionString(, , "tbl_UserlogIN")
    'Call ChangeConnectionString(, , "vw_RecentEvents")
    Call LogUsersin
    Call CreateShortCutMenus
    'Call CustomMenus ' not needed
    Call MinMaxForm
    HidShw = "Hide"
    Me.HideShow = "Hide"
    On Error GoTo MoveOn
    If WhereX Like "*Local*" Then Me.DevLabel.caption = "Office"
    If WhereX Like "*RDS*" Then
       Me.DevLabel.caption = "Remote"
       Call SetAdobePDFAsDefault
    End If
    If WhereX Like "*OutSide*" Then Me.DevLabel.caption = "OutSide"
    If WhereX Like "*Phoenix*" Then Me.DevLabel.caption = "Phoenix"
    If WhereX Like "*Phoenix*" Then Me.FormHeader.BackColor = RGB(204, 102, 0) Else Me.FormHeader.BackColor = RGB(0, 0, 0)

    'If DevOrPro = "Dev" Then Me.DevLabel.Visible = True Else Me.DevLabel.Visible = False
    ''get Screensize from functions
    '' 2/2/2021 change to fixed location
    ScnX = GetSystemMetrics(0)
    If ScnX > 1920 Then TwipX = fPixelsToTwips(1920, DIRECTION_HORIZONTAL)
    If ScnX <= 1920 Then TwipX = fPixelsToTwips(1550, DIRECTION_HORIZONTAL)
    'TwipX = fPixelsToTwips(GetSystemMetrics(0), DIRECTION_HORIZONTAL)
    TwipY = fPixelsToTwips(GetSystemMetrics(1), DIRECTION_VERTICAL)
    ' DoCmd.MoveSize (TwipX - (Me.Form.Width + 200)), 0
MoveOn:
    If DLookup("[Active User]", "TBL_Employees", "[User ID] = '" & LCase(Environ("username")) & "'") = True Then
        GoTo AccessGranted
    Else
        MsgBox "You are not authorized to use Atlaas Core at this time. Please contact technical support to gain access", vbCritical, "LOGIN"
        DoCmd.CloseDatabase
    End If
AccessGranted:
    Me.UserName = FAUserFullName
    Me.USERID = FaUserName & " (" & FAUserAccessLevel & ")"
    Me.Version = AtlaasVersion
    Me.Version.ControlTipText = CurrentDb.Name
    'check if user restricted
    If FAUserAccessLevel = 0 Then
        MsgBox "You are not authorized to use Atlaas Core at this time. Please contact technical support to gain access", vbCritical, "LOGIN"
        DoCmd.CloseDatabase
    End If
    '  Call Form_Timer
    Call CheckSpecialUser
    'If Me.DevLabel.caption = "Office" And FAUserName = "ljohnson" Then Call AddReference("Selenium", "C:\Program Files\SeleniumBasic\Selenium32.tlb", "add")
    If Me.DevLabel.caption = "Office" Or Me.DevLabel.caption = "Remote" Then Call AddReference("Selenium", "C:\Program Files\SeleniumBasic\Selenium32.tlb", "add")

    '  Call HideShowEvents(0)
    '  Call HideShowEvents(101)

Dim scriptPath As String
Dim command As String
    scriptPath = """C:\Program Files\SeleniumBasic\UpdateEdgeDriver.ps1"""
    command = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"" -ExecutionPolicy Bypass -File " & scriptPath
    Shell command, vbHide

FixMe:
End Sub
Private Sub CheckSpecialUser()
    If FAUserAccess = 99 Then    'Or FAUserIDx = 27 Then
        Me.ScheduleControl.Width = 1920
        Me.TreeCommand.Visible = True
    Else
        Me.ScheduleControl.Width = 3195
        Me.TreeCommand.Visible = False
    End If
    DoCmd.MoveSize , , , 9000
End Sub

Private Sub Form_Timer()
    'If CInt(Left(Time(), 1)) >= 3 And Right(Time(), 2) = "PM" And CheckIfSendBLCDEmailSent = False Then Call BLCDSyndFinalMessage(0)
    'Me.EventList.Requery
End Sub

Private Sub HideShow_Click()
    If Me.HideShow = "Hide" Then Call HideShowEvents(0) Else Call HideShowEvents(101)
End Sub
Public Sub HideShowEvents(NumEvents As Integer)
    'Select Case NumEvents
        ''If HidShw = "Hide" Then HidShw = "Show" Else HidShw = "Hide"
    'Case 0
    '    NumEvents = 0
    '    HidShw = "Show"
    '    Me.EventList.Height = 4
    '    DoCmd.MoveSize , , , 7300
    'Case 101
    '    NumEvents = 10
    '    HidShw = "Hide"
    '    Me.EventList.Height = 4500
    '    DoCmd.MoveSize , , , 9500
    'Case Else
    '    HidShw = "Hide"
    '    Me.EventList.Height = 4500
    '    DoCmd.MoveSize , , , 9500
    'End Select
    'SQLz = "SELECT  DISTINCTROW TOP " & NumEvents & " [EventDate] & ' ' & [EventDesc] AS [Desc], MLID, SchID, SGrpID, UnitID, NoticeDate, EventType " & _
           "FROM vw_RecentEvents ORDER BY NoticeDate DESC, SGrpID;"
    'Me.EventList.RowSource = SQLz
    'Me.HideShow = HidShw
    'Me.USERID.SetFocus
End Sub
Private Sub MinMaxForm()
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

Private Sub MinMax_Click()
    Call MinMaxForm
End Sub

Private Sub S10_Click()
  '  Call HideShowEvents(10)
End Sub
Private Sub S25_Click()
 '   Call HideShowEvents(25)
End Sub
Private Sub S50_Click()
'    Call HideShowEvents(50)
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

Private Sub LinkedDocManager_Click()
    Call OpenTarget("LinkedDocManager")
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
    Dim grpAccess As String
    grpAccess = FAUserAccessGroup
    If grpAccess Like "*SYND*" Or grpAccess Like "*ITT*" Or grpAccess Like "*APT*" Or grpAccess Like "*TMT*" Or grpAccess Like "*ACCT*" Then
        Call OpenTarget("ManagementMenu")
    Else
        Call MsgBox("Sorry You Do Not Have Accesss", vbCritical, "Slow down partner")
    End If
End Sub
Private Sub Operations_Click()
    Call OpenTarget("OPS Main Menu")
End Sub

Private Sub Search_Clients_Click()
    Call OpenTarget("SignOffDashboard")
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
    Call LogUsersout
        Application.SetOption "Auto compact", False
        If Format(Date, "DDD") Like "Thu" Then Application.SetOption "Auto compact", True
        Call AddReference("Selenium", "C:\Program Files\SeleniumBasic\Selenium32.tlb", "Delete")
        DoCmd.Quit
    End If
End Sub

Private Sub CustomMenus()
    Dim Leveli As Integer
    Leveli = FAUserAccess
    If Leveli = 2 Then DoCmd.OpenForm "LO Menu"
    If Leveli = 3 Then DoCmd.OpenForm "AcctMenu"
    If Leveli = 4 Then DoCmd.OpenForm "AM_Menu"
    If Leveli = 5 Then DoCmd.OpenForm "DocsMenu"
    If Leveli = 6 Then DoCmd.OpenForm "ManagementMenu"
    If Leveli = 7 Then DoCmd.OpenForm "OPS VP Menu"
    If Leveli = 8 Then DoCmd.OpenForm "Marketing VP Menu"

    If Leveli = 10 Then DoCmd.OpenForm "Rmkt Menu"
    If Leveli = 11 Then DoCmd.OpenForm "ManagementMenuPF"
End Sub
