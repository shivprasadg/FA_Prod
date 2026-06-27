VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindoSizeMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Compare Database
'Dim SQL, SQLz, SearchText, UPDwn As String
'Dim x As Integer, y As Integer, cx As Integer, cy As Integer
'
'Dim RST As DAO.Recordset
'
'Private Sub Form_Open(Cancel As Integer)
'Me.USERID = "ljohnson" 'GetUserName
''Me.UserId = Environ("USERNAME")  'GetUserName
'SearchText = "Type 'Menu Item' Search Here"
'Me.VersionNumber = "RELEASE 1.99 REVISION 4.3"
'Me.Form.Requery
'End Sub
'
'Private Sub Form_Load()
'Dim td As TableDef
'Dim stConnect As String
''    For Each td In CurrentDb.TableDefs
''    If td.SourceTableName Like "dbo.*" Then
''    td.RefreshLink
''    'Debug.Print td.SourceTableName ' Kevin is THE Man!
''    End If
''Next
'
''Call CheckVersionNumber(Me.VersionNumber)
'On Error GoTo MoveOn
'Dim PiX, TwipX  As Integer ' X = HORIZONTAL
'Dim PiY, TwipY  As Integer ' Y = VERTICAL
''get Screensize from functions
'TwipX = fPixelsToTwips(GetSystemMetrics(0), DIRECTION_HORIZONTAL)
'TwipY = fPixelsToTwips(GetSystemMetrics(1), DIRECTION_VERTICAL)
'DoCmd.MoveSize (TwipX - (Me.Form.Width + 200)), 0
'Call Min_Max_Click
'MoveOn:
'HideShow = "Show" ' Set MinMax status
'Call CueSearchBoxes("ItemSearch", SearchText)
'Call SetupMenus
'If Me.UserType <> "MGR" Then Me.Min_Max.Visible = False
'UPDwn = "Up"
'
'    'DoCmd.RunCommand acCmdAppMinimize
'Me.Form.Requery
'Me.Form.Refresh
'End Sub
'Private Sub SetupMenus()
' SQLz = "SELECT UserReportID AS Item, UserReportCmd, UserPermit, UserAppl " & _
'        "FROM TblAccessUsersOptions " & _
'        "WHERE UserWindowId = '" & Me.USERID & "' " & _
'        "GROUP BY UserReportID, UserReportCmd, UserPermit, UserAppl " & _
'        "ORDER BY TblAccessUsersOptions.UserReportID, TblAccessUsersOptions.UserAppl;"
'Me.Items.RowSource = SQLz
' SQLz = "SELECT ReportID AS [Table (Excel)], ReportCmd " & _
'        "From TblAccessReports " & _
'        "GROUP BY ReportID, ReportCmd, GroupName, ReportOwn " & _
'        "HAVING ReportOwn = '" & Me.USERID & "' " & _
'        "ORDER BY ReportID;"
'Me.Pivots.RowSource = SQLz
' SQLz = "SELECT GroupName AS [Group], UserAppl From TblAccessUsersOptions " & _
'        "WHERE UserWindowId = '" & Me.USERID & "' GROUP BY GroupName, UserAppl " & _
'        "HAVING GroupName <> '' ORDER BY GroupName;"
'Me.GroupList.RowSource = SQLz
'
'       SQL = "DELETE sys_RecentForms.* " & _
'             "From sys_RecentForms WHERE DateOpened <> #" & Date & "#;"
'       DoCmd.SetWarnings False
'         DoCmd.RunSQL (SQL)
'         DoCmd.SetWarnings True
'End Sub
'
'Private Sub Form_Timer()
'Me.TimeX.Requery
'End Sub
'
'Private Sub GroupList_Click()
'Call CueSearchBoxes("ItemSearch", SearchText)
'SQLz = "SELECT UserReportID AS Item, UserReportCmd, UserPermit, UserReportStatus AS Type, UserAppl " & _
'"From TblAccessUsersOptions " & _
'"WHERE GroupName = '" & Me.GroupList & "' And UserWindowId = '" & Me.USERID & "' " & _
'"GROUP BY UserReportID, UserReportCmd, UserPermit, UserReportStatus, UserAppl " & _
'"ORDER BY UserAppl;"
'Me.Items.RowSource = SQLz
' SQLz = "SELECT ReportID AS [Table (Excel)], ReportCmd " & _
'        "From TblAccessReports " & _
'        "GROUP BY ReportID, ReportCmd, GroupName, ReportOwn " & _
'        "HAVING GroupName = '" & Me.GroupList & "' And  ReportOwn = '" & Me.USERID & "' " & _
'        "ORDER BY ReportID;"
'Me.Pivots.RowSource = SQLz
'End Sub
'
'Private Sub Items_DblClick(Cancel As Integer)
'On Error GoTo Fixme
'If Me.Items.Column(1) Like "Call*" Then
'   Call UpdateMe
'Else
'  DoCmd.close acForm, Me.Items.Column(1)
'  DoCmd.OpenForm Me.Items.Column(1)
'   Call UpdateOpened(Me.Items.Column(0), Me.Items.Column(1))
'Exit Sub
'End If
'Fixme:
'MsgBox err.Description, vbCritical, "TDM Error"
'End Sub
'
'Private Sub ItemSearch_Change()
'Me.GroupList = Null
'SQLz = "SELECT UserReportID AS Item, UserReportCmd, UserPermit, UserAppl " & _
'       "FROM TblAccessUsersOptions " & _
'       "WHERE UserWindowId='" & Me.USERID & "' AND UserReportID Like '*" & Me.itemSearch.Text & "*' " & _
'       "GROUP BY UserReportID, UserReportCmd, UserPermit, UserAppl " & _
'       "ORDER BY UserAppl;"
'Me.Items.RowSource = SQLz
' SQLz = "SELECT ReportID AS [Table (Excel)], ReportCmd " & _
'        "From TblAccessReports " & _
'        "GROUP BY ReportID, ReportCmd, GroupName, ReportOwn " & _
'        "HAVING ReportOwn = '" & Me.USERID & "' AND ReportID Like '*" & Me.itemSearch.Text & "*' " & _
'        "ORDER BY ReportID;"
'Me.Pivots.RowSource = SQLz
'End Sub
'
'Private Sub ItemSearch_GotFocus()
'Call ActvSeachBoxes(Me.ActiveControl.Name, Nz(Me.ActiveControl, ""))
'End Sub
'Private Sub ItemSearch_LostFocus()
' If Me.itemSearch.Value & "Test" = "Test" Then Call CueSearchBoxes("ItemSearch", SearchText)
'End Sub
''-----------------------------------
''set boxes
''-----------------------------------
'Private Sub CueSearchBoxes(ByVal FieldBoX As String, ByVal Text2Display As String)
'Me(FieldBoX).FontItalic = True
'Me(FieldBoX).ForeColor = 9145219
'Me(FieldBoX).Value = Text2Display
'End Sub
'Private Sub ActvSeachBoxes(ByVal FieldBoX As String, ByVal Text2Display As String)
'Me(FieldBoX).FontItalic = False
'Me(FieldBoX).ForeColor = 0
'If Text2Display Like "Type*" Or Text2Display = PcText Then Me(FieldBoX).Value = ""
'End Sub
'
'Private Sub UpdateOpened(FormX As String, CommandX As String)
'SQL = "SELECT * From sys_RecentForms;"
'Set RST = CurrentDb.OpenRecordset(SQL, dbOpenSnapshot)
'If RST.EOF Then GoTo Post:
'RST.MoveFirst
'   Do Until RST.EOF
'   If RST!formname = FormX Then GoTo Exithere
'   RST.MoveNext
'   Loop
'Post:
'SQL = "INSERT INTO sys_RecentForms ( FormName, Command, DateOpened ) " & _
'"SELECT '" & FormX & "' AS FormName, '" & [CommandX] & "' AS CommandX, #" & Date & "# AS DateOpened;"
'DoCmd.SetWarnings False
'DoCmd.RunSQL (SQL)
'DoCmd.SetWarnings True
'Exithere:
'Me.OpenFormList.Requery
'Set RST = Nothing
'End Sub
'Private Sub ClearRecent_Click()
'    SQL = "DELETE sys_RecentForms.* " & _
'         "From sys_RecentForms;"
'         DoCmd.SetWarnings False
'    DoCmd.RunSQL (SQL)
'         DoCmd.SetWarnings True
'    Me.OpenFormList.Requery
'End Sub
'Private Sub OpenFormList_DblClick(Cancel As Integer)
''DoCmd.Close acForm, Me.OpenFormList.Column(1)
'DoCmd.OpenForm Me.OpenFormList.Column(1)
'End Sub
'
'Private Sub Pivots_DblClick(Cancel As Integer)
'On Error GoTo FixPivot
'Dim oApp As Object
'Dim PivotPath As String
' PivotPath = Pivots.Column(1)
'Set oApp = CreateObject("Excel.Application")
'oApp.Visible = True
'With oApp
'Visible = True
'oApp.Workbooks.Open PivotPath
'End With
'Exit Sub
'FixPivot:
'MsgBox "There was an error" & vbNewLine & " The File " & PivotPath & " is either currupt or it cannot be found.", , "File Not There"
'End Sub
'Private Sub Min_Max_Click()
'If UPDwn = "Down" Then
'x = 50
'y = 100
'cx = 1800
'cy = 1550
'xg_SizeWindow "Access", x, y, cx, cy
'CommandBars.ActiveMenuBar.Enabled = True
'    DoCmd.SelectObject acTable, "CustomerSPX", True
'    'DoCmd.RunCommand acCmdWindowUnhide
'UPDwn = "Up"
'Else
'x = 1600
'y = 100
'cx = 75
'cy = 75
'xg_SizeWindow "Access", x, y, cx, cy
'CommandBars.ActiveMenuBar.Enabled = False
'    DoCmd.SelectObject acTable, "CustomerSPX", True
'    DoCmd.RunCommand acCmdWindowHide
'
'UPDwn = "Down"
'End If
'End Sub
'Private Sub Restore_Click()
'Dim FrmNameX(25) As String
'Dim FormX As Form
'Dim OpenForms As Object
'Dim i As Integer
'i = 1
'Set OpenForms = Application.Forms
'For Each FormX In OpenForms
'FrmNameX(i) = FormX.Name
'i = i + 1
'Next FormX
'For i = 1 To Forms.Count
'DoCmd.OpenForm FrmNameX(i)
'Next i
'Call GoHomeForm(Me.Form.Name)
'HideShow = "Show"
'End Sub
'Private Sub HideForms_Click()
'DoCmd.RunCommand acCmdAppRestore
'DoCmd.RunCommand acCmdAppMinimize
'Me.Min_Max.caption = "Show DB"
'Call GoHomeForm(Me.Form.Name)
'End Sub
'Private Sub Close_all_Click()
'If Forms.Count = 1 Then Exit Sub
'Dim Response, FormNameX As String
'Dim Cancel As Boolean
'FormNameX = Me.Form.Name
'Response = MsgBox("Are you sure you want close all Opened Forms?", vbQuestion + vbYesNo + vbCritical, "Close All?")
'If Response = vbYes Then
'Do While Forms.Count > 0
'DoCmd.close acForm, Forms(0).Name
'Loop
'Else
'Cancel = True
'End If
'Call GoHomeForm(FormNameX)
'End Sub
'Private Sub exitdb_Click()
'Application.SetOption "Auto compact", False
'Dim Response, FormNameX As String
'Dim Cancel As Boolean
'FormNameX = Me.Form.Name
'Response = MsgBox("Are you sure you want to EXIT?", vbQuestion + vbYesNo + vbCritical, "Close TDMs?")
'If Response = vbYes Then
'   If Me.CompactMe = -1 Then
'      Application.SetOption "Auto compact", True
'      MsgBox "TDMs wil compact itself, it will take a few moments.", , "Compacting"
'   End If
'   DoCmd.Quit
'End If
'End Sub
'
'
'
'Private Sub UpdateTDM_Click()
'Call UpdateMe
'End Sub
'
'Public Sub VersionNumber_DblClick(Cancel As Integer)
'DoCmd.close acForm, "sys_RevCurrent"
'DoCmd.OpenForm "sys_RevCurrent"
'End Sub
'Private Sub Form_Close()
'
'End Sub
