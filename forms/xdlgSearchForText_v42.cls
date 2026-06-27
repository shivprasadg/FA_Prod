Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'This code can be used in part or its entirety in any application, but please reference me,
'Brent Spaulding and my call sign 'datAdrenaline', along with in the code or in an acknowledgement form.
'Also, please contact me via datAdrenaline[the at symbol]aol.com or on www.utteraccess.com as user datAdrenaline.
'if you have any questions or comments regarding the utility.

#If VBA7 Then
    Private Declare PtrSafe Sub psSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub psSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Dim blStopSearch As Boolean

Private Sub btnAbout_Click()
'Activate the 'About' tab-page
    Me.pgAbout.SetFocus
    Me.drpSearchText.SetFocus
End Sub

Private Sub btnClearHistory_Click()
'Clear the drop down search history
    Me.drpSearchText.RowSource = ""
    Me.drpSearchText.SetFocus
End Sub

Private Sub btnClearInputs_Click()
'Clear ALL the text boxes

    Me.drpSearchText = ""
    Me.drpSearchText.SetFocus
    Me.txtQueryResults = ""
    Me.txtFormResults = ""
    Me.txtReportResults = ""
    Me.txtTableResults = ""
    Me.txtModuleResults = ""
    Me.txtMacroResults = ""
    Me.chkQueries = False
    Me.chkForms = False
    Me.chkReports = False
    Me.chkTables = False
    Me.chkModules = False
    Me.chkMacros = False
    Me.chkCBars = False
    Me.chkSkipReplacement = True

    Me.txtWhereFound = ""
    Me.txtValue = ""

    Me.tabResultsPages.Value = Me.pgQueryResults.PageIndex

End Sub

Private Sub btnStartSearch_Click()
'Start the search process

    'Enumerate objects to search
    Const conNothing = 0
    Const conQueries = 1
    Const conForms = 2
    Const conReports = 3
    Const conTables = 4
    Const conModules = 5
    Const conMacros = 6
    Const conCBars = 7

    Dim byteReturnFocusTo As Byte

    'Make sure something is in the "Search For" text box, if not exit the sub
    If Nz(Me.drpSearchText, "") = "" Then
        MsgBox "You need to have something to search for!", vbExclamation
        Me.drpSearchText.SetFocus
        Exit Sub
    End If

    'Make sure at least one set of objects is set to be searched
    If Not Me.chkForms _
       And Not Me.chkQueries _
       And Not Me.chkReports _
       And Not Me.chkTables _
       And Not Me.chkModules _
       And Not Me.chkMacros _
       And Not Me.chkCBars Then

        MsgBox "You need to select at least one set of objects to search!", vbExclamation
        Me.drpSearchText.SetFocus
        Exit Sub
    End If

    'Initialize some stuff
    blStopSearch = False
    byteReturnFocusTo = 0
    Me.btnStopSearch.Enabled = True
    Me.btnStopSearch.SetFocus
    Me.btnStartSearch.Enabled = False
    Me.btnClearInputs.Enabled = False
    Me.btnClearHistory.Enabled = False
    Me.btnAbout.Enabled = False


    'Search the selected type of object
    If Me.chkQueries = True Then
        SearchQueries
        If byteReturnFocusTo = conNothing Then byteReturnFocusTo = conQueries
    End If

    If Me.chkForms = True Then
        SearchForms
        If byteReturnFocusTo = conNothing Then byteReturnFocusTo = conForms
    End If

    If Me.chkReports = True Then
        SearchReports
        If byteReturnFocusTo = conNothing Then byteReturnFocusTo = conReports
    End If

    If Me.chkTables = True Then
        SearchTables
        If byteReturnFocusTo = conNothing Then byteReturnFocusTo = conTables
    End If

    If Me.chkModules = True Then
        SearchModules
        If byteReturnFocusTo = conNothing Then byteReturnFocusTo = conModules
    End If

    If Me.chkMacros = True Then
        SearchMacros
        If byteReturnFocusTo = conNothing Then byteReturnFocusTo = conMacros
    End If

    If Me.chkCBars = True Then
        SearchCommandBars
        If byteReturnFocusTo = conNothing Then byteReturnFocusTo = conCBars
    End If

    'Re-enable buttons
    Me.btnStartSearch.Enabled = True
    Me.btnStartSearch.SetFocus
    Me.btnStopSearch.Enabled = False
    Me.btnClearInputs.Enabled = True
    Me.btnClearHistory.Enabled = True
    Me.btnAbout.Enabled = True

    'Return the focus
    Me.SetFocus
    Select Case byteReturnFocusTo
        Case conQueries
            Me.tabResultsPages.Value = Me.pgQueryResults.PageIndex
        Case conForms
            Me.tabResultsPages.Value = Me.pgFormResults.PageIndex
        Case conReports
            Me.tabResultsPages.Value = Me.pgReportResults.PageIndex
        Case conTables
            Me.tabResultsPages.Value = Me.pgTableResults.PageIndex
        Case conModules
            Me.tabResultsPages.Value = Me.pgModuleResults.PageIndex
        Case conMacros
            Me.tabResultsPages.Value = Me.pgMacroResults.PageIndex
        Case conCBars
            Me.tabResultsPages.Value = Me.pgC_Bars.PageIndex
    End Select

    'Reset the replacement page
    Me.txtWhereFound = ""
    Me.txtValue = ""

    'Remember the history of the searches performed (only for the current session
    'Closing the form also clears the history).
    Me.drpSearchText.RowSource = Me.drpSearchText & ";" & Me.drpSearchText.RowSource

End Sub


Private Sub SearchForms()
'Search Forms

    Dim dbs As DAO.DataBase
    Dim cnt As DAO.Container
    Dim doc As DAO.Document

    Dim frm As Form
    Dim ctl As Control
    Dim prp As DAO.Property
    Dim pg As Page
    Dim pgprp As DAO.Property
    Dim pgctl As Control
    Dim pgctlprp As DAO.Property

    Dim strObjects As String
    Dim strReplMark As String
    Dim x As Integer

    'Initialize some stuff
    strObjects = ""
    Me.txtFormResults = ""
    Me.tabResultsPages.Value = Me.pgFormResults.PageIndex

    Set dbs = CodeDb
    Set cnt = dbs.Containers!Forms

    'Run time error handling
    On Error Resume Next

    'Loop through each form in the current database
    x = 1 'x is the form number being worked on
    For Each doc In cnt.Documents

        'Check to see if search should be stopped
        Me.pgFormResults.SetFocus
        If blStopSearch Then
            Exit For
        End If

        'Continue search
        If doc.Name <> Me.Name Then  'Don't check this form since it is not really a part of the app

            'Provide progress to user
            Me.txtFormResults = "Working on form " & x & " of " & cnt.Documents.Count & " (" & doc.Name & ")"
            Me.Repaint

            'Open the form
            Echo False  'Turn screen display off
            DoCmd.OpenForm doc.Name, acDesign, , , , acHidden
            Set frm = Forms(doc.Name)

            'Enumerate the form properties
            For Each prp In frm.Properties
                If InStr(prp.Value, Me.drpSearchText) > 0 Then
                    If Err = 0 Then
                        strReplMark = ""
                        If Me.chkSkipReplacement = False Then
                            Echo True
                            psReplacement "F: " & frm.Name & " : " & prp.Name, prp.Value
                            Echo False
                            If Me.chkReplace = True Then
                                Me.chkReplace = False
                                prp.Value = Me.txtValue
                                If Err.Number <> 0 Then
                                    MsgBox Err.Description
                                    Err.Clear
                                Else
                                    strReplMark = "(r)"
                                End If
                            End If
                        End If
                        strObjects = strObjects & vbCrLf & strReplMark & frm.Name & " : " & prp.Name
                    Else
                        Err.Clear
                    End If
                End If
            Next prp

            'Enumerate the controls on the form
            For Each ctl In frm.Controls
                For Each prp In ctl.Properties
                    If InStr(prp.Value, Me.drpSearchText) > 0 Then
                        If Err = 0 Then
                            strReplMark = ""
                            If Me.chkSkipReplacement = False Then
                                Echo True
                                psReplacement "F: " & frm.Name & " : " & ctl.Name & " : " & prp.Name, prp.Value
                                Echo False
                                If Me.chkReplace = True Then
                                    Me.chkReplace = False
                                    prp.Value = Me.txtValue
                                    If Err.Number <> 0 Then
                                        MsgBox Err.Description
                                        Err.Clear
                                    Else
                                        strReplMark = "(r)"
                                    End If
                                End If
                            End If
                            strObjects = strObjects & vbCrLf & strReplMark & frm.Name & " : " & ctl.Name & " : " & prp.Name
                        Else
                            Err.Clear
                        End If
                    End If
                Next prp

                'Enumerate sub-sontrols if the current control is a tab control
                If ctl.ControlType = acTabCtl Then
                    'Analyze the tab contol pages
                    For Each pg In ctl.Pages

                        'Page properties
                        For Each pgprp In pg.Properties
                            If InStr(pgprp.Value, Me.drpSearchText) > 0 Then
                                If Err = 0 Then
                                    strReplMark = ""
                                    If Me.chkSkipReplacement = False Then
                                        Echo True
                                        psReplacement "F: " & frm.Name & " : " & pg.Name & " : " & pgprp.Name, pgprp.Value
                                        Echo False
                                        If Me.chkReplace = True Then
                                            Me.chkReplace = False
                                            pgprp.Value = Me.txtValue
                                            If Err.Number <> 0 Then
                                                MsgBox Err.Description
                                                Err.Clear
                                            Else
                                                strReplMark = "(r)"
                                            End If
                                        End If
                                    End If
                                    strObjects = strObjects & vbCrLf & frm.Name & " : " & pg.Name & " : " & pgprp.Name
                                Else
                                    Err.Clear
                                End If
                            End If
                        Next pgprp

                        'Controls on the page
                        For Each pgctl In pg.Controls
                            For Each pgctlprp In pgctl.Properties
                                If InStr(pgctlprp.Value, Me.drpSearchText) > 0 Then
                                    If Err = 0 Then
                                        strReplMark = ""
                                        If Me.chkSkipReplacement = False Then
                                            Echo True
                                            psReplacement "F: " & frm.Name & " : " & ctl.Name & " : " & pg.Name & " : " & pgctl.Name & " : " & pgctlprp.Name, pgctlprp.Value
                                            Echo False
                                            If Me.chkReplace = True Then
                                                Me.chkReplace = False
                                                pgctlprp.Value = Me.txtValue
                                                If Err.Number <> 0 Then
                                                    MsgBox Err.Description
                                                    Err.Clear
                                                Else
                                                    strReplMark = "(r)"
                                                End If
                                            End If
                                        End If
                                        strObjects = strObjects & vbCrLf & strReplMark & frm.Name & " : " & ctl.Name & " : " & pg.Name & " : " & pgctl.Name & " : " & pgctlprp.Name
                                    Else
                                        Err.Clear
                                    End If
                                End If
                            Next pgctlprp
                        Next pgctl
                    Next pg
                End If
            Next ctl

            'Close the form
            DoCmd.close acForm, frm.Name, acSavePrompt

            Echo True  'Turn on the screen
        End If

        'Increment the forms counter
        x = x + 1

        'Process pending events (catches the 'Stop Search' button click)
        DoEvents

    Next doc

    'Print the results to the window
    Me.txtFormResults = "Forms Found:" & vbCrLf

    If strObjects <> "" Then
        Me.txtFormResults = Me.txtFormResults & strObjects
    Else
        Me.txtFormResults = Me.txtFormResults & "None"
    End If

    'Clean up objects
    Set dbs = Nothing
    Set cnt = Nothing
    Set doc = Nothing

    Set frm = Nothing
    Set ctl = Nothing
    Set prp = Nothing
    Set pg = Nothing
    Set pgprp = Nothing
    Set pgctl = Nothing
    Set pgctlprp = Nothing

End Sub


Private Sub SearchModules()
'Search all the modules in the database for the text

    Dim dbs As DAO.DataBase
    Dim cnt As DAO.Container
    Dim doc As DAO.Document

    Dim mdl As Module
    Dim frm As Form
    Dim rpt As Report

    Dim blCanSearchModule As Boolean
    Dim strModules As String
    Dim intItem As Integer
    Dim intDocumentCount As Integer
    Dim intModuleType As Integer
    Dim strModuleType As String
    Dim lngStartLine As Long
    Dim lngStartCol As Long
    Dim lngEndLine As Long
    Dim lngEndCol As Long
    Dim lngProcKind As Long
    Dim lngOccurrences As Long
    Dim strProcNameCurrent As String
    Dim strProcNameOld As String
    Dim strTemp As String

    Const conUnattachedModules = 1
    Const conFormModules = 2
    Const conReportModules = 3

    'Initialize some stuff
    strModules = ""
    Me.txtModuleResults = ""
    blCanSearchModule = False
    Me.tabResultsPages.Value = Me.pgModuleResults.PageIndex
    Set dbs = CodeDb

    'Run time error handling
    On Error Resume Next

    'Loop through each module type in the current database
    For intModuleType = conUnattachedModules To conReportModules

        'Check to see if search should be stopped
        If blStopSearch Then
            Exit For
        End If

        'Get the right container for the type of modules being examined
        Select Case intModuleType
        Case conUnattachedModules
            Set cnt = dbs.Containers!Modules
            cnt.Documents.Refresh
            strModuleType = "unattached"
            intDocumentCount = cnt.Documents.Count
        Case conFormModules
            Set cnt = dbs.Containers!Forms
            cnt.Documents.Refresh
            strModuleType = "form"
            intDocumentCount = cnt.Documents.Count
        Case conReportModules
            Set cnt = dbs.Containers!Reports
            cnt.Documents.Refresh
            strModuleType = "report"
            intDocumentCount = cnt.Documents.Count
        End Select


        'Loop through the modules of the current type
        intItem = 1

        For Each doc In cnt.Documents

            'Check to see if search should be stopped
            If blStopSearch Then
                Exit For
            End If

            'Initialization for this this block of modules
            Echo False
            Select Case intModuleType

            Case conUnattachedModules
                blCanSearchModule = True
                DoCmd.OpenModule doc.Name
                Set mdl = Modules(doc.Name)

            Case conFormModules
                blCanSearchModule = False
                If doc.Name <> Me.Name And doc.Name <> "frmKeepAlive" Then
                    DoCmd.OpenForm doc.Name, acDesign
                    Set frm = Forms(doc.Name)
                    If frm.HasModule = True Then
                        Set mdl = frm.Module
                        blCanSearchModule = True
                    End If
                End If

            Case conReportModules
                blCanSearchModule = False
                DoCmd.OpenReport doc.Name, acViewDesign
                Set rpt = Reports(doc.Name)
                If rpt.HasModule = True Then
                    Set mdl = rpt.Module
                    blCanSearchModule = True
                End If

            End Select

            lngStartLine = 0
            lngStartCol = 0
            lngEndLine = 0
            lngEndCol = 0
            lngOccurrences = 0
            strProcNameOld = ""
            strProcNameCurrent = ""

            'Search through the module then close it
            If blCanSearchModule Then

                'Provide feedback to user
                Me.txtModuleResults = "Working on " & strModuleType & " module " & intItem & " of " & intDocumentCount & " (" & doc.Name & ")"
                Me.Repaint

                'Loop through all the occurrences of the search text
                Do Until Not mdl.Find(Me.drpSearchText, _
                                      lngStartLine, lngStartCol, _
                                      lngEndLine, lngEndCol)

                    'Increment the occurrence counter
                    lngOccurrences = lngOccurrences + 1

                    'Record the name of the current proc (or if it is not in a proc)
                    strProcNameCurrent = mdl.ProcOfLine(lngStartLine, lngProcKind)
                    If strProcNameCurrent = "" Then
                        strProcNameCurrent = "(Declarations)"
                    End If

                    'Determine if in a different proc
                    If strProcNameCurrent = strProcNameOld Or strProcNameOld = "" Then
                        strProcNameOld = strProcNameCurrent
                        strProcNameCurrent = ""
                        strTemp = mdl.Name & ": " & strProcNameOld & _
                                        ": Occurrences " & lngOccurrences

                    Else
                        strTemp = mdl.Name & ": " & strProcNameOld & _
                                        ": Occurrences " & lngOccurrences - 1

                        'Record the previous procs occurrences
                        strModules = strModules & vbCrLf & strTemp

                        'Set up for the new proc
                        strProcNameOld = strProcNameCurrent
                        strProcNameCurrent = ""
                        lngOccurrences = 1
                        strTemp = mdl.Name & ": " & strProcNameOld & _
                                        ": Occurrences " & lngOccurrences
                    End If

                    'Increment search zones
                    lngStartLine = lngEndLine
                    lngStartCol = lngEndCol
                    lngEndLine = 0
                    lngEndCol = 0

                Loop

                'Record occurrences on the last proc or the only proc
                If strProcNameCurrent = "" And Len(strProcNameOld) > 0 And lngOccurrences > 0 Then
                    strModules = strModules & vbCrLf & strTemp
                End If

            Else '(Cant search the module)

                'Provide feedback to user
                Me.txtModuleResults = "Skipping " & strModuleType & " module " & intItem & " of " & intDocumentCount & " (" & doc.Name & ")"
                Me.Repaint

            End If

            'Release the mdl object
            Set mdl = Nothing

            'Close the module (and parents)
            Select Case intModuleType
            Case conUnattachedModules
                DoCmd.close acModule, doc.Name

            Case conFormModules
                Set frm = Nothing
                If Me.Name <> doc.Name And doc.Name <> "frmKeepAlive" Then
                    DoCmd.close acForm, doc.Name
                End If

            Case conReportModules
                Set rpt = Nothing
                DoCmd.close acReport, doc.Name

            End Select

            'Increment the item number
            intItem = intItem + 1

            'Process pending events
            DoEvents
            Echo True
        Next doc
    Next intModuleType

    'Print the results to the window
    Me.txtModuleResults = "Modules Found:" & vbCrLf

    If strModules <> "" Then
        Me.txtModuleResults = Me.txtModuleResults & strModules
    Else
        Me.txtModuleResults = Me.txtModuleResults & "None"
    End If

    'Clean up objects
    Set dbs = Nothing
    Set cnt = Nothing
    Set doc = Nothing
    Set mdl = Nothing
    Set frm = Nothing
    Set rpt = Nothing
    If Err.Number <> 0 Then
        Err.Clear
    End If
    Echo True

End Sub


Private Sub SearchQueries()
'Search Queries for a text string

    Dim qry As DAO.QueryDef
    Dim prp As DAO.Property
    Dim strQueries As String
    Dim lngCharWhere As Long
    Dim x As Integer, Y As Integer
    Dim strReplMark As String

    'Initialized some stuff
    Y = 1  'y is a counter for which query is being looked at
    Me.txtQueryResults = ""
    Me.tabResultsPages.Value = Me.pgQueryResults.PageIndex

    'Run-time error handling
    On Error Resume Next

    'Loop through each query in the current database
    For Each qry In CodeDb.QueryDefs

        'Check to see if search is to be halted, Else continue
        Me.pgQueryResults.SetFocus
        If blStopSearch Then
            Exit For
        End If

        'Search through the properties (unless the name begins with '~' - a temp query)
        If Left(qry.Name, 1) <> "~" Then

            'Indicate which query is being analyzed
            Me.txtQueryResults = "Working on query " & Y & " of " & CodeDb.QueryDefs.Count & " (" & qry.Name & ")"
            Me.Repaint

            'Search the query properties (except the SQL)
            For Each prp In qry.Properties
                If prp.Name <> "SQL" Then
                    If InStr(prp.Value, Me.drpSearchText) > 0 Then
                        If Err = 0 Then
                            strReplMark = ""
                            If Me.chkSkipReplacement = False Then
                                psReplacement "Q: " & qry.Name & " : " & prp.Name, prp.Value
                                If Me.chkReplace = True Then
                                    Me.chkReplace = False
                                    prp.Value = Me.txtValue
                                    If Err.Number <> 0 Then
                                        MsgBox Err.Description
                                        Err.Clear
                                    Else
                                        strReplMark = "(r)"
                                    End If
                                End If
                            End If
                            strQueries = strQueries & vbCrLf & strReplMark & qry.Name & " : " & prp.Name
                        Else
                            Err.Clear
                        End If
                    End If
                End If
            Next prp


            'Search the query's SQL statement
            x = 1  'x is the character position in the SQL statement
            Do
                lngCharWhere = Nz(InStr(x, qry.SQL, Me.drpSearchText), 0)
                If lngCharWhere > 0 Then
                    x = x + lngCharWhere
                    strQueries = strQueries & vbCrLf & qry.Name & ": At SQL Char " & x - 1
                End If
            Loop Until lngCharWhere = 0

            If x > 1 Then
                strReplMark = ""
                If Me.chkSkipReplacement = False Then
                    psReplacement "Q: " & qry.Name & " : SQL (" & x & ")", qry.SQL
                    If Me.chkReplace = True Then
                        Me.chkReplace = False
                        qry.SQL = Me.txtValue
                        If Err.Number <> 0 Then
                            MsgBox Err.Description
                            Err.Clear
                        Else
                            strReplMark = "(r)"
                            strQueries = strQueries & vbCrLf & strReplMark & qry.Name & " : SQL"
                            strReplMark = ""
                        End If
                    End If
                End If
            End If

        Else

            'Indicate which query is being skipped (temp queries created by Access)
            Me.txtQueryResults = "Skipping query " & Y & " of " & CodeDb.QueryDefs.Count & " (" & qry.Name & ")"
            Me.Repaint

        End If

        Y = Y + 1  'Increment the query count

        'Process pending events (catches the 'Stop Search' button click)
        DoEvents

    Next qry


    'Print the results in the window
    Me.txtQueryResults = "Queries Found: " & vbCrLf

    If strQueries <> "" Then
        Me.txtQueryResults = Me.txtQueryResults & strQueries
    Else
        Me.txtQueryResults = Me.txtQueryResults & "None Found"
    End If

    'Clean up objects
    Set qry = Nothing

End Sub



Private Sub SearchReports()
'Search Reports.  This procedure is a pretty much a mirror of 'SearchForms' except references to forms
'were changed to reflect Reports.  i.e.:
'       SearchForms                         SearchReports
'       cnt = dbs.Containers!Forms  <-->    cnt = dbs.Containers!Reports
'       frm.<Something>             <-->    rpt.<Something>
'
'Also a loop to check the reports' Grouping and Sorting was added

    Dim dbs As DAO.DataBase
    Dim cnt As DAO.Container
    Dim doc As DAO.Document

    Dim rpt As Report
    Dim ctl As Control
    Dim prp As DAO.Property
    Dim pg As Page
    Dim pgprp As DAO.Property
    Dim pgctl As Control
    Dim pgctlprp As DAO.Property

    Dim strObjects As String

    Dim x As Integer
    Dim Y As Integer
    Dim z As Integer
    Dim Txt As String
    Dim strReplMark As String

    'Initialize some stuff
    strObjects = ""
    Me.txtReportResults = ""
    Me.tabResultsPages.Value = Me.pgReportResults.PageIndex

    Set dbs = CodeDb
    Set cnt = dbs.Containers!Reports

    'Run-time error handlding
    On Error Resume Next

    'Loop through each report in the current database
    z = 1 'z is the number of the report being looked at
    For Each doc In cnt.Documents

        'Check to see if seach should stop, else continue
        Me.pgReportResults.SetFocus
        If blStopSearch Then
            Exit For
        End If

        'Provide progress to user
        Me.txtReportResults = "Working on report " & z & " of " & cnt.Documents.Count & " (" & doc.Name & ")"
        Me.Repaint

        'Open the report
        Echo False
        DoCmd.OpenReport doc.Name, acDesign
        Set rpt = Reports(doc.Name)

        'Enumerate the report properties
        For Each prp In rpt.Properties
            If InStr(prp.Value, Me.drpSearchText) > 0 Then
                If Err = 0 Then
                    strReplMark = ""
                    If Me.chkSkipReplacement = False Then
                        Echo True
                        psReplacement "R: " & rpt.Name & " : " & prp.Name, prp.Value
                        Echo False
                        If Me.chkReplace = True Then
                            Me.chkReplace = False
                            prp.Value = Me.txtValue
                            If Err.Number <> 0 Then
                                MsgBox Err.Description
                                Err.Clear
                            Else
                                strReplMark = "(r)"
                            End If
                        End If
                    End If
                    strObjects = strObjects & vbCrLf & strReplMark & rpt.Name & " : " & prp.Name
                Else
                    Err.Clear
                End If
            End If
        Next prp

        'Analyze the reports' Grouping/Sorting settings

        'Count the number of groups/sorting (UBound doesn't work)
        x = 0 'Holds a psuedo-count of the number of grouping levels in the report
        Do Until Err <> 0 Or x = 15 '(Access 97 only allows for 10 Grouping/Sorting settings, so this just in case ... I dont want an endless loop!)
            Txt = rpt.GroupLevel(x).ControlSource
            x = x + 1
        Loop
        If Err <> 0 Then
            Err.Clear
        End If

        Y = x - 2  'The max number for an index value of the GroupLevel array

        'Loop through the grouping and sorting
        For x = 0 To Y
            If Not IsError(rpt.GroupLevel(x).ControlSource) And InStr(rpt.GroupLevel(x).ControlSource, Me.drpSearchText) > 0 Then
                If Err = 0 Then
                    strReplMark = ""
                    If Me.chkSkipReplacement = False Then
                        Echo True
                        psReplacement "R: " & rpt.Name & " : Grouping and Sorting : GroupLevel(" & x & ") : ControlSource", rpt.GroupLevel(x).ControlSource
                        Echo False
                        If Me.chkReplace = True Then
                            Me.chkReplace = False
                            rpt.GroupLevel(x).ControlSource = Me.txtValue
                            If Err.Number <> 0 Then
                                MsgBox Err.Description
                                Err.Clear
                            Else
                                strReplMark = "(r)"
                            End If
                        End If
                    End If
                    strObjects = strObjects & vbCrLf & strReplMark & rpt.Name & " : Grouping and Sorting : GroupLevel(" & x & ") : ControlSource"
                Else
                    Err.Clear
                End If
            End If
        Next x

        'Enumerate the controls on the report
        For Each ctl In rpt.Controls
            For Each prp In ctl.Properties
                If Not IsError(prp.Value) And InStr(prp.Value, Me.drpSearchText) > 0 Then
                    If Err = 0 Then
                        strReplMark = ""
                        If Me.chkSkipReplacement = False Then
                            Echo True
                            psReplacement "R: " & rpt.Name & " : " & ctl.Name & " : " & prp.Name, prp.Value
                            Echo False
                            If Me.chkReplace = True Then
                                Me.chkReplace = False
                                prp.Value = Me.txtValue
                                If Err.Number <> 0 Then
                                    MsgBox Err.Description
                                    Err.Clear
                                Else
                                    strReplMark = "(r)"
                                End If
                            End If
                        End If
                        strObjects = strObjects & vbCrLf & strReplMark & rpt.Name & " : " & ctl.Name & " : " & prp.Name
                    Else
                        Err.Clear
                    End If
                End If
            Next prp

            'Enumerate sub-controls if the current control is a tab control
            If ctl.ControlType = acTabCtl Then

                'Loop through each page in the tab control
                For Each pg In ctl.Pages

                    'Page properties
                    For Each pgprp In pg.Properties
                        If Not IsError(pgprp.Value) And InStr(pgprp.Value, Me.drpSearchText) > 0 Then
                            If Err = 0 Then
                                strReplMark = ""
                                If Me.chkSkipReplacement = False Then
                                    Echo True
                                    psReplacement "R: " & rpt.Name & " : " & pg.Name & " : " & pgprp.Name, pgprp.Value
                                    Echo False
                                    If Me.chkReplace = True Then
                                        Me.chkReplace = False
                                        pgprp.Value = Me.txtValue
                                        If Err.Number <> 0 Then
                                            MsgBox Err.Description
                                            Err.Clear
                                        Else
                                            strReplMark = "(r)"
                                        End If
                                    End If
                                End If
                                strObjects = strObjects & vbCrLf & strReplMark & rpt.Name & " : " & pg.Name & " : " & pgprp.Name
                            Else
                                Err.Clear
                            End If
                        End If
                    Next pgprp

                    'Page controls
                    For Each pgctl In pg.Controls
                        For Each pgctlprp In pgctl.Properties
                            If Not IsError(pgctlprp.Value) And InStr(pgctlprp.Value, Me.drpSearchText) > 0 Then
                                If Err = 0 Then
                                    strReplMark = ""
                                    If Me.chkSkipReplacement = False Then
                                        Echo True
                                        psReplacement "R: " & rpt.Name & " : " & pgctl.Name & " : " & pgctlprp.Name, pgctlprp.Value
                                        Echo False
                                        If Me.chkReplace = True Then
                                            Me.chkReplace = False
                                            pgctlprp.Value = Me.txtValue
                                            If Err.Number <> 0 Then
                                                MsgBox Err.Description
                                                Err.Clear
                                            Else
                                                strReplMark = "(r)"
                                            End If
                                        End If
                                    End If
                                    strObjects = strObjects & vbCrLf & strReplMark & rpt.Name & " : " & ctl.Name & " : " & pg.Name & " : " & pgctl.Name & " : " & pgctlprp.Name
                                Else
                                    Err.Clear
                                End If
                            End If
                        Next pgctlprp
                    Next pgctl
                Next pg
            End If
        Next ctl

        'Close the report
        DoCmd.close acReport, rpt.Name, acSavePrompt
        Echo True

        'Increment report counter
        z = z + 1

        'Process pending events (catches the "Stop Search" button click)
        DoEvents

    Next doc

    'Print the results to the window
    Me.txtReportResults = "Reports Found:" & vbCrLf

    If strObjects <> "" Then
        Me.txtReportResults = Me.txtReportResults & strObjects & vbCrLf
    Else
        Me.txtReportResults = Me.txtReportResults & "None"
    End If

    'Clean up objects
    Set dbs = Nothing
    Set cnt = Nothing
    Set doc = Nothing

    Set rpt = Nothing
    Set ctl = Nothing
    Set prp = Nothing
    Set pg = Nothing
    Set pgprp = Nothing
    Set pgctl = Nothing
    Set pgctlprp = Nothing

End Sub

Private Sub SearchMacros()
'This sub will save a macro as text, then search that text file for the string pattern specified.
'The temp file created will be deleted.

    Dim db As DAO.DataBase
    Dim ctr As DAO.Container
    Dim doc As DAO.Document
    Dim strMacros As String
    Dim strTempFile As String

    Set db = CodeDb
    Set ctr = db.Containers!Scripts

    For Each doc In ctr.Documents

        If Left(doc.Name, 4) <> "~TMP" Then

            'Save the macro as a text file
            strTempFile = Environ("TEMP") & "\~" & doc.Name & ".mcr"
            SaveAsText acMacro, doc.Name, strTempFile

            'Load that text file into a variable
            Dim intChannel As Integer
            intChannel = FreeFile

            Open strTempFile For Input Access Read As #intChannel

            Dim strMacroText As String
            strMacroText = ""

            Do Until EOF(intChannel)

                Dim strLine As String

                'Read in a line from the file
                Line Input #intChannel, strLine

                'Exit the reading if the line starts all the XML definition.
                If Trim(strLine) Like "Comment =""_AXL:<?xml version=*" Then _
                    Exit Do

                'Append the line to the macro
                strMacroText = strMacroText & strLine

            Loop

            'Close the file
            Close #intChannel

            'Count the occurances of the specified string pattern
            Dim lngOccurrances As Long
            lngOccurrances = UBound(Split(strMacroText, Me.drpSearchText, , vbTextCompare))

            If lngOccurrances > 0 Then _
                strMacros = strMacros & doc.Name & " (" & lngOccurrances & " occurrances.)" & vbCrLf

            'Delete the file
            Kill strTempFile

        End If

    Next 'doc (macro)

    'Print the results to the window
    Me.txtMacroResults = "Macros Found:" & vbCrLf

    If strMacros <> "" Then
        Me.txtMacroResults = Me.txtMacroResults & strMacros
    Else
        Me.txtMacroResults = Me.txtMacroResults & "None"
    End If

    'Clean up
    Set doc = Nothing
    Set ctr = Nothing
    Set db = Nothing

End Sub


Public Sub SearchTables()
'Search through table definitions

    Dim dbs As DAO.DataBase
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim prp As DAO.Property
    Dim strField As String
    Dim x As Integer

    'Initialize some stuff
    Set dbs = CodeDb
    Me.txtTableResults = ""
    Me.tabResultsPages.Value = Me.pgTableResults.PageIndex

    'Run-time error handling
    On Error Resume Next

    x = 1 'x contains the form number being worked on
    DoCmd.Hourglass True  'Turn the hourglass on

    'Loop through each table
    For Each tdf In dbs.TableDefs

        'Check to see if search sould be stopped, else continue
        If blStopSearch Then
            Exit For
        End If

        'Check to see if current table is a temp or system table
        If Left(tdf.Name, 1) <> "~" And Left(tdf.Name, 4) <> "MSys" Then

            'Print progress for the user
            Me.txtTableResults = "Working on table " & x & " of " & dbs.TableDefs.Count & " (" & tdf.Name & ")"
            Me.Repaint

            'Analyse the table properties
            For Each prp In tdf.Properties
                If InStr(prp.Name, Me.drpSearchText) > 0 Then
                    If Err = 0 Then
                        strField = strField & vbCrLf & tdf.Name & " : " & prp.Name
                    Else
                        Err.Clear
                    End If
                End If
            Next prp

            'Analyze each field and its properties in the table
            For Each fld In tdf.fields
                For Each prp In fld.Properties
                    If InStr(prp.Value, Me.drpSearchText) > 0 Then
                        If Err = 0 Then
                            strField = strField & vbCrLf & tdf.Name & " : " & fld.Name & " : " & prp.Name
                        Else
                            Err.Clear
                        End If
                    End If
                Next prp
            Next fld

        Else

            'Print progress for the user for tables being skipped (temp tables from Access97)
            Me.txtTableResults = "Skipping table " & x & " of " & dbs.TableDefs.Count & " (" & tdf.Name & ")"
            Me.Repaint

        End If

        'Increment the table count
        x = x + 1

        'Process pending events (catches the 'stop search' button click)
        DoEvents

    Next tdf

    DoCmd.Hourglass False 'Turn hourglass off


    'Print Results
    Me.txtTableResults = "Tables Found:" & vbCrLf

    If strField <> "" Then
        Me.txtTableResults = Me.txtTableResults & strField
    Else
        Me.txtTableResults = Me.txtTableResults & "None"
    End If

    'Clean up objects
    Set dbs = Nothing
    Set tdf = Nothing
    Set fld = Nothing
    Set prp = Nothing

End Sub


Private Function SearchCommandBarControls(cbr As Variant) As String '(cbr As Office.CommandBar) As String

    Dim cbc As Variant 'Office.CommandBarControl
    Dim strCommands As String
    Dim intProperty As Integer
    Dim blVisible As Boolean

    ' Check name of CBar
    If InStr(cbr.Name, Me.drpSearchText) > 0 Then
        ' Have a match
        strCommands = strCommands & cbr.Name & vbCrLf
    End If

    ' Loop thru' controls collection checking each for text
    For Each cbc In cbr.Controls
        ' Visibilty seems to be an issue _
            store status for restore then make vis.

        blVisible = cbc.Visible
        cbc.Visible = True

        ' Some controls have embedded CBars _
            i.e.  The Popup with nested (sub) menu items

        If cbc.Type = 10 Then 'msoControlPopup
            ' Make a recursive call to get sub items
            Set cbr = cbc.CommandBar
            strCommands = strCommands & SearchCommandBarControls(cbr)
        Else

            If InStr(cbc.caption, Me.drpSearchText) > 0 Then
                ' Have a match in Caption
                strCommands = strCommands & cbr.Name & ": " & cbc.caption & ": Caption " & vbCrLf
            End If

            If InStr(cbc.DescriptionText, Me.drpSearchText) > 0 Then
                ' Have a match in Description
                strCommands = strCommands & cbr.Name & ": " & cbc.caption & ": DescriptionText" & ": " & vbCrLf
            End If

            If InStr(cbc.OnAction, Me.drpSearchText) > 0 Then
                ' Have a match in OnAction
                strCommands = strCommands & cbr.Name & ": " & cbc.caption & ": OnAction" & vbCrLf
            End If

            If InStr(cbc.Parameter, Me.drpSearchText) > 0 Then
                ' Have a match in Parameter
                strCommands = strCommands & cbr.Name & ": " & cbc.caption & ": Parameter" & vbCrLf
            End If

            If InStr(cbc.Tag, Me.drpSearchText) > 0 Then
                ' Have a match in Tag
                strCommands = strCommands & cbr.Name & ": " & cbc.caption & ": Tag " & vbCrLf
            End If

            If InStr(cbc.TooltipText, Me.drpSearchText) > 0 Then
                ' Have a match in Tooltip
                strCommands = strCommands & cbr.Name & ": " & cbc.caption & ": TooltipText" & vbCrLf
            End If
        End If

        ' Restore visibility
        cbc.Visible = blVisible
    Next cbc

    SearchCommandBarControls = strCommands
End Function


Private Sub SearchCommandBars()

    Dim cbr As Variant 'Office.CommandBar
    Dim cbc As Variant 'Office.CommandBarControl
    Dim strCommands As String

    Dim lngOccurrences As Long

    Dim x As Integer

    Me.txtC_BarResults = ""
    Me.tabResultsPages.Value = Me.pgC_Bars.PageIndex

    'Run-time error handling
    On Error Resume Next

    x = 1 'x contains the form number being worked on
    DoCmd.Hourglass True  'Turn the hourglass on

    For Each cbr In CommandBars

        'Check to see if search sould be stopped, else continue
        If blStopSearch Then
            Exit For
        End If

        'Print progress for the user
        Me.txtC_BarResults = "Working on commandbar " & x & " of " & CommandBars.Count & " (" & cbr.Name & ")"
        Me.Repaint

        If Not cbr.BuiltIn Then
            strCommands = strCommands + SearchCommandBarControls(cbr)
        End If

        ' Increment the Command Bar count
        x = x + 1

        'Process pending events (catches the 'stop search' button click)
        DoEvents

    Next cbr

    'Print Results
    Me.txtC_BarResults = "CommandBars Found:" & vbCrLf & vbCrLf

    If strCommands <> "" Then
        Me.txtC_BarResults = Me.txtC_BarResults & strCommands
    Else
        Me.txtC_BarResults = Me.txtC_BarResults & "None"
    End If

    DoCmd.Hourglass False 'Turn hourglass off
End Sub



Private Sub btnStopSearch_Click()
    blStopSearch = True
End Sub


Private Sub Form_Open(Cancel As Integer)

    Dim ref As Reference    'A library reference
    Dim blReferenceFound As Boolean  'a flag indicating the success/failure of finding a reference
    Dim intResponse As String 'a user response to a message
    Dim strDAOFullPath As String 'The file path and name of the dao reference

    Const constrDAO = "DAO"

    'Find the DAO reference
    For Each ref In Application.References
        If ref.Name = constrDAO Then
            blReferenceFound = True
            Exit For
        End If
    Next ref

    'Notify about missing reference to DAO
    If Not blReferenceFound Then

        intResponse = MsgBox("Search for Text was unable to find a reference to the DAO Object library." & _
                   vbCrLf & vbCrLf & "This library is needed to run this form!" & _
                   vbCrLf & vbCrLf & "The DAO object library is identified as:" & _
                   vbCrLf & " 'Microsoft Office xx.x Access database engine objects' for v2007 and higher" & _
                   vbCrLf & " 'Microsoft DAO 3.xx Object Library' for v2003 and lower" & _
                   vbCrLf & vbCrLf & "Please reference the appropriate library, then recompile your VBA project." & _
                   vbCrLf & vbCrLf & vbCrLf & " ... do you want to open the form anyway?", vbCritical + vbYesNo)

        If intResponse = vbNo Then
            Cancel = True
        Else
            'Do nothing
        End If

    End If
End Sub


Private Sub psReplacement(strWhereFound As String, strFoundIn As String)

    Me.txtWhereFound = strWhereFound
    Me.txtValue = strFoundIn

    Me.SetFocus
    Me.pgReplace.SetFocus

    Do Until Me.chkContinue = True Or blStopSearch = True
        DoEvents
        psSleep 100
    Loop

    Me.chkContinue = False

End Sub
