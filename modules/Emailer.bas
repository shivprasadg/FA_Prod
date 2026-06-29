Attribute VB_Name = "Emailer"
Option Compare Database
Public Mailto As Variant
Public MailFrom As String
Public MailCc As Variant
Public Mailsubject As String
Public msg1 As String
Public regard As String
Public Attachpath(1) As Variant


Public Function EmailTask(UserX As String, MailToX As Variant, MailFromX As String, MailsubjectX As String, TaskIDx As String, _
                          DTSubmitX As String, TaskPriorityX As String, StatusX As String)
    Dim dbs As DataBase
    Dim SQL As String
    Dim Dpath As String
    Forms!Taskmaker.DTSubmitted = Now()
    Forms!Taskmaker.TaskNotes = Forms!Taskmaker.TaskNotes & vbNewLine & ">Sent:" & Now & "<" & vbNewLine
    Forms!Taskmaker.TaskStatus = "Sent."
    'make email message
    msg1 = msg1 & "<html><body><h2>Fleet Advanatge / Request Form </h2><br/><br/>"
    msg1 = msg1 & "RE:" & TaskIDx & " : " & MailsubjectX & "Priority: " & TaskPriorityX & "<br/><br/>"
    msg1 = msg1 & "Dear IT Department,<br/><br/><br/>"
    msg1 = msg1 & "Task Notes:<br/>"
    msg1 = msg1 & "<b>" & Forms!Taskmaker.TaskNotes & "</b><br/><br/>"
    msg1 = msg1 & "<br /><br />Submitted " & Forms!Taskmaker.DTSubmitted & "<br /><br />"
    msg1 = msg1 & "Thank You,<br /><br />" & UserX & "</body></html>"

    Call SendEmailMessage(MailToX, MailsubjectX, msg1, False, True, False)

End Function

Public Function SendEmailMessage(varTo As Variant, ByVal strSubject As String, ByVal strBody As String, _
                            bolAutoSend As Boolean, bolSaveInOutbox As Boolean, bolAddSignature As Boolean, _
                            Optional varCC As Variant, Optional varBCC As Variant, Optional varReplyTo As Variant, _
                            Optional varAttachmentPath As Variant, Optional varImagePath As Variant, Optional varHtmlFooter As Variant) As Boolean
'=================================================================
'
'varto: a string of email addresses, multiples delimted by semi-colon
'strSubject: subject line for the email
'strBody: body of the email, must be wrapped in <html> </html> tags, and optionally any other formatting tags
'bolAutoSend: determines whether email is sent automatically or displayed to the user
'bolSaveInOutbox: determines if the message is saved in the outbox
'boladdsignature: determines if the user's default signature is added to the outgoing email
'varCC: (Optional) CC email addresses, multiples delimited by semi-colon
'varBCC: (Optional) BCC email addresses, multiples delimited by semi-colon
'varReplyTo (Optional) If specified sets the reply to email address, multiples delimited by semi-colon
'varAttachmentPath: (Optional) If specified attaches the file
'varImagePath: (Optional) If specified embeds the image in the body of the email
'varHtmlFooter: (Optional) If specifed, inserts an html file as a footer to the message.
'ASSUMPTIONS: Outlook, HTML message format, Word is the default editor
'When performing some of the optional steps the message is constructed in the following order
'signature then embedded image then footer then body text, so the actual email would read as follows
'body text, footer, embedded image, signature
'=================================================================

    Dim i As Integer
    Dim strTempFile As String
    Dim strMsg As String
    Dim objOutlook As Outlook.Application
    Dim objOutlookMsg As Outlook.MailItem
    Dim objOutlookRecip As Outlook.Recipient
    Dim objOutlookAttach As Outlook.Attachment
    Dim objInsp As Outlook.Inspector
    Dim objWord As Word.Application
    Dim objdoc As Word.Document
    Dim objrange As Word.Range
    Dim PathToOutlook As String

    ' check if Outlook is Open and if 32 or 64 bit path
    If fIsAppRunning("Outlook") = False Then
        PathToOutlook = Split(Environ("path"), ";")(0) & "outlook.exe"
        RetVal = Shell(PathToOutlook, vbMaximizedFocus)
        GoTo SentEmail
    End If

SentEmail:
    SendEmailMessage = False
  On Error GoTo outlookopen:    'Trap if OutLook is NOT running

    Set objOutlook = New Outlook.Application                                'Create the Outlook session.
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)                   'Create the message.

On Error GoTo HandleError    ' Check for other Errors
    strBody = ReplaceCRLFwithBR(strBody)                                    'Replace any vbcrlf with <br>

    If (InStr(strBody, "<font") = 0) Or (InStr(strBody, "<html>") = 0) Then                              'if no <html> and <font> tags then wrap the body of the message with these tags
        strBody = FormatAsHtml(strBody)
    End If

    With objOutlookMsg

        If Not IsMissing(varTo) Then
            If varTo <> "" And Not IsNull(varTo) Then
                For i = 1 To CountCSWords(varTo)
                    Set objOutlookRecip = .recipients.Add(GetCSWord(varTo, i))               'Add the TO recipient(s) to the message.
                    objOutlookRecip.Type = olTo
                Next i
            End If
        End If

        On Error GoTo HandleError    ' Check for other Errors

        If Not IsMissing(varCC) Then
            If varCC <> "" And Not IsNull(varCC) Then
                For i = 1 To CountCSWords(varCC)
                    Set objOutlookRecip = .recipients.Add(GetCSWord(varCC, i))                  'Add the cc recipient(s) to the message.
                    objOutlookRecip.Type = olCC
                Next i
            End If
        End If

        If Not IsMissing(varBCC) Then
            If varBCC <> "" And Not IsNull(varBCC) Then
                For i = 1 To CountCSWords(varBCC)
                    Set objOutlookRecip = .recipients.Add(GetCSWord(varBCC, i))                 'Add the bcc recipient(s) to the message.
                    objOutlookRecip.Type = olBCC
                Next i
            End If
        End If

        If Not IsMissing(varReplyTo) Then
            If varReplyTo <> "" And Not IsNull(varReplyTo) Then
                For i = 1 To CountCSWords(varReplyTo)
                    Set objOutlookRecip = .ReplyRecipients.Add(GetCSWord(varReplyTo, i))        'Add the bcc recipient(s) to the message.
                Next i
            End If
        End If

        ' Add attachments to the message.
        If Not IsMissing(varAttachmentPath) Then
             If IsArray(varAttachmentPath) Then
                For i = LBound(varAttachmentPath) To UBound(varAttachmentPath) - 1
                    If varAttachmentPath(i) <> "" And varAttachmentPath(i) <> "False" Then
                        Set objOutlookAttach = .Attachments.Add(varAttachmentPath(i))
                    End If
                Next i
            Else
                'If varAttachmentPath <> "" And varAttachmentPath(i) <> "False" Then
                Set objOutlookAttach = .Attachments.Add(varAttachmentPath)
                'End If
            End If
        End If

        ' end new routine
        .subject = strSubject               'Set the Subject of the message.

        .BodyFormat = olFormatHTML          'set format to html

        If bolAddSignature Or Not IsMissing(varImagePath) Or Not IsMissing(varHtmlFooter) Then  'if signature or embedded image or html footer
            Set objInsp = objOutlookMsg.GetInspector                                            'this causes the default signature to be added to the message
            Set objdoc = objInsp.WordEditor
            ' If objdoc Is Nothing Then
            '        strmsg = "Outlook must use Word as the email editor. Follow these instructions to fix the problem." & vbCrLf & vbCrLf & _
                     '           "Tools>Options" & vbCrLf & "Then select 'Mail Format' tab" & vbCrLf & "Ensure Use Microsoft Office Word 2003 to edit e-mail messages."
            '      MsgBox strmsg
            '     objOutlookMsg.Close olDiscard
            '    GoTo SendMessage_Done
            'End If

            Set objWord = objdoc.Application

            If bolAddSignature = False Then         'If the user had a signature it would have been applied, if we dont want it then we need to delete it here
                objdoc.Range.Delete
            End If

            If Not IsMissing(varImagePath) Then
                If varImagePath <> "" And Not IsNull(varImagePath) Then
                    If Dir(varImagePath, vbHidden + vbSystem + vbReadOnly + vbDirectory) <> "" Then
                        On Error Resume Next
                        .Display                                                                        'Seems like word document must be visible before you can use addpicture method
                        If Err <> 0 Then            'if the mail cound not be displayed then display a warning and discard the message
                            MsgBox "It was not possible to display the message, check that there are no dialog boxes open in Outlook." & vbCrLf & "Please close all Outlook windows and emails, and then attempt this update again.", vbCritical
                            .close olDiscard
                            GoTo SendMessage_Done
                        End If
                        objWord.WindowState = wdWindowStateMinimize                                     'minimize word application so user does not see mail being created
                        Set objrange = objdoc.Range(start:=0, End:=0)                                   'goto start of message again
                        objrange.InsertBefore vbCrLf
                        objdoc.InlineShapes.AddPicture fileName:=varImagePath, LinkToFile:=False, SaveWithDocument:=True, Range:=objrange
                    End If
                End If
            End If

            If Not IsMissing(varHtmlFooter) Then
                If varHtmlFooter <> "" And Not IsNull(varHtmlFooter) Then
                    If Dir(varHtmlFooter, vbHidden + vbSystem + vbReadOnly + vbDirectory) <> "" Then
                        Set objrange = objdoc.Range(start:=0, End:=0)                                   'goto start of message
                        objrange.InsertFile varHtmlFooter, , , False, False   'insert the html from the external file
                    End If
                End If
            End If

            strTempFile = Environ("temp") & Format(Now(), "yyyymmddhhnnss") & ".htm"        'generate temp filename
            Set objrange = objdoc.Range(start:=0, End:=0)                                   'goto start of message again
            CreateTextFile strTempFile, strBody                                             'save the bodytext as a temporary htm file
            objrange.InsertFile strTempFile, , , False, False                               'insert the htm file into the body of the message
            Kill strTempFile                                                                'delete temp file

            objdoc.SpellingChecked = True                                                   'doesnt matter for autosend, but helps the user if the message is being displayed
        Else
            .HTMLBody = strBody
        End If

        If bolSaveInOutbox = False Then             'if message not to be saved after sending
            .DeleteAfterSubmit = True               'specify that it should be deleted
        End If

        If (bolAutoSend = True) And (.recipients.Count > 0) Then        'check that there is at least 1 recipient before trying to autosend
            .Send
        Else
            Err = 0
            On Error Resume Next
            .Display                    'Attempt to display the message
            If Err <> 0 Then            'if the mail cound not be displayed then display a warning and discard the message
                MsgBox "It was not possible to display the message, check that there are no dialog boxes open in Outlook." & vbCrLf & "Please close all Outlook windows and emails, and then attempt this update again.", vbCritical
                .close olDiscard
                GoTo SendMessage_Done
            End If
        End If
    End With

    SendEmailMessage = True

SendMessage_Done:
    Set objOutlook = Nothing
    Set objOutlookMsg = Nothing
    Set objOutlookRecip = Nothing
    Set objOutlookAttach = Nothing
    Set objInsp = Nothing
    Set objWord = Nothing
    Set objdoc = Nothing
    Set objrange = Nothing
    Exit Function

HandleError:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume SendMessage_Done
    Exit Function
outlookopen:        'Exit if Outlook is NOT runnning
    MsgBox "Sorry, Can't Send Email - You Must Have Outlook Open", vbCritical, "OutLook Not Open"
    Exit Function
End Function

Public Function ReplaceCRLFwithBR(ByVal strText) As String
'=================================================================
'Replace all vbcrlf with <br> to keep line breaks in html emails
'=================================================================

    strText = Replace(strText, Chr(13), "<br>")
    strText = Replace(strText, Chr(10), "")

    ReplaceCRLFwithBR = strText


End Function

Public Function FormatAsHtml(ByVal str) As String
'=================================================================
'Wraps a string in html tags
'=================================================================

    FormatAsHtml = "<html><font face=""arial"" size=""2"">" & str & "</font></html>"

End Function

Public Function GetCSWord(ByVal s, INDX As Integer, Optional StrDelimiter = ";") As String
'=================================================================
'Returns the nth word in a specific field
'=================================================================

    On Error Resume Next

    GetCSWord = Split(s, StrDelimiter)(INDX - 1)

End Function

Function CountCSWords(ByVal str, Optional StrDelimiter = ";") As Integer
'=================================================================
'Counts the words in the delimited string
'=================================================================

    Dim Wc As Integer, pos As Integer

    If VarType(str) <> 8 Or Len(str) = 0 Then
        CountCSWords = 0
        Exit Function
    End If
    Wc = 1
    pos = InStr(str, StrDelimiter)
    Do While pos > 0
        Wc = Wc + 1
        pos = InStr(pos + 1, str, StrDelimiter)
    Loop

    CountCSWords = Wc

End Function


Public Sub CreateTextFile(strFullPath As String, strText As String)
'=================================================================
'Creates a text file with the specified file name containing the supplied text
'=================================================================

    Dim fso As Object
    Dim MyFile As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set MyFile = fso.CreateTextFile(strFullPath, True)          'Creates file, existing file will be overwritten
    MyFile.WriteLine (strText)                                  'writes string to the file
    MyFile.close                                                'close the file


End Sub

Public Function GetTextFile(ByVal strFile As String) As String
'=================================================================
'Returns a string that contains the contents of the specified file
'=================================================================

    Dim fso As Object
    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(strFile).OpenAsTextStream(1, -2)       'Open the file ForReading, Use system default for file format
    GetTextFile = ts.readall                                    'Read the contents of the file
    ts.close                                                    'Close the file


End Function

Function GetBoiler(ByVal sFile As String) As String
'Dick Kusleika - This gets the signature from the current user for emails
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.close
End Function

Public Function ValidEmail(pAddress As String) As Boolean
    Dim oRegEx As Object
    Set oRegEx = CreateObject("VBScript.RegExp")
    With oRegEx
        .Pattern = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
        ValidEmail = .Test(pAddress)
    End With
    Set oRegEx = Nothing
End Function

Public Function Testemail()
    Dim testx As String
    Dim ReplyX As Long
    testx = "12234"
    IntStringCount testx, "2", ReplyX
    Dim rX As String
    Call PassEmailAddress(Emailaddr, rX)
    MsgBox rX & vbNewLine & ReplyX
End Function

Public Function IntStringCount(ByVal vInWhat As String, ByVal vCountWhat As String, _
                               ByRef rOccurs As Long, Optional ByVal vStart As Long)

    vStart = InStr(vStart + 1, vInWhat, vCountWhat)
    If (vStart = 0) Then
        Exit Function
    Else
        rOccurs = rOccurs + 1
        IntStringCount vInWhat, vCountWhat, rOccurs, vStart
    End If
End Function

Public Function PassEmailAddress(ByVal strEmail As String, _
                                 Optional ByRef strReason As String) As String
    On Error GoTo Err_Handler

    Dim strPrefix As String
    Dim strSuffix As String
    Dim strMiddle As String
    Dim lngCharacter As Long
    Dim strBuffer As String
    Dim blnContinue As String
    Dim strMsg As String

    blnContinue = True
    'default = true

    strEmail = Trim(strEmail)

    If Len(strEmail) < 8 Then
        strReason = "Too short for a valid email address."
        blnContinue = False
    End If

    If blnContinue = True Then
        strMsg = CheckAt(strEmail)
        If Len(strMsg) > 0 Then
            blnContinue = False
            strReason = strMsg
        End If
    End If


    PassEmailAddress = strReason


    'Dropped through to here, so email address OK

Exit_Handler:
    Exit Function
Err_Handler:
    PassEmailAddress = False
    Debug.Print Err.Number, Err.Description, Now
    Resume Exit_Handler

End Function

Private Function CheckAt(strTxtIn As String) As String
    On Error GoTo Err_Handler

    Dim lngCharacter As Long
    Dim strMsg As String
    Dim strBuffer As String
    Dim strBuffer2 As String
    Dim strReason As String
    Dim blnContinue As String

    blnContinue = True
    'default = true

    strBuffer = strTxtIn
    If InStr(strBuffer, "@") = 0 Then
        strReason = "Missing the @ needed in a valid email address."
        blnContinue = False
    End If

    If blnContinue = True Then
        If InStr(InStr(strBuffer, "@") + 1, strBuffer, "@") < 0 Then
            strReason = "Too many @ for a valid email address"
            blnContinue = False
        End If
    End If

    If blnContinue = True Then
        If InStr(strBuffer, ".") = 0 Then
            strReason = "Missing the period needed in a valid email address."
            blnContinue = False
        End If
    End If

    If blnContinue = True Then
        If InStr(strBuffer, "@") = 1 Or InStr(strBuffer, "@") = Len(strBuffer) Or _
           InStr(strBuffer, ".") = 1 Or InStr(strBuffer, ".") = Len(strBuffer) Then
            strReason = "Not a valid format for an email address."
            blnContinue = False
        End If
    End If

    On Error Resume Next
    If blnContinue = True Then
        strBuffer = Right(strBuffer, 4)
        If InStr(strBuffer, ".") = 0 Then
            strMsg = TooLong(strBuffer)
            If Len(strMsg) > 0 Then
                blnContinue = False
                strReason = strMsg
            End If
        End If
    End If

    If blnContinue = True Then
        If Left(strBuffer, 1) = "." Then strBuffer = Right(strBuffer, 3)
        If Left(Right(strBuffer, 3), 1) = "." Then strBuffer = Right(strBuffer, 2)
        If Left(Right(strBuffer, 2), 1) = "." Then strBuffer = Right(strBuffer, 1)

        If Len(strBuffer) < 2 Then
            strReason = "Suffix (ending) too short for a valid email address."
            blnContinue = False
        End If
    End If

    CheckAt = strReason

Exit_Handler:
    Exit Function
Err_Handler:
    Debug.Print Err.Number, Err.Description, Now
    Resume Exit_Handler
End Function

Private Function TooLong(strTxtIn As String) As String
    On Error GoTo Err_Handler

    Dim strReason As String

    If Len(strTxtIn) > 3 And Len(strTxtIn) <> "mobi" Then

        strReason = "Suffix (ending) too long for a valid email address."
        TooLong = strReason
    End If

Exit_Handler:
    Exit Function
Err_Handler:
    Debug.Print Err.Number, Err.Description, Now
    Resume Exit_Handler
End Function

Public Function SendPowerShellEmail(Emailto, subjectx, bodyx As String, Optional EmailCC As String, Optional wAttachedFile As String)
    Dim Ps1    As String
    Ps1 = Ps1 & "$EmailFrom = ""Atlaas_Core@FleetAdvantage.com""" & vbNewLine
    Ps1 = Ps1 & "$Passwd = ConvertTo-SecureString 'Cloud2017' -AsPlainText -Force" & vbNewLine
    Ps1 = Ps1 & "$Credentials  = New-Object System.Management.Automation.PSCredential($EmailFrom, $Passwd)" & vbNewLine
    Ps1 = Ps1 & "$EmailTo = " & Emailto & vbNewLine
    If EmailCC <> "" Then Ps1 = Ps1 & "$EmailCC = " & EmailCC & vbNewLine
    If wAttachedFile <> "" Then Ps1 = Ps1 & "$Attachments = """ & wAttachedFile & """" & vbNewLine
    Ps1 = Ps1 & "$Subject =""" & subjectx & """" & vbNewLine
    Ps1 = Ps1 & "$Body =""" & bodyx & """" & vbNewLine
    Ps1 = Ps1 & "$SMTPServer = ""smtp.office365.com""" & vbNewLine

    Ps1 = Ps1 & "Send-MailMessage "
    Ps1 = Ps1 & "-To $EmailTo "
    Ps1 = Ps1 & "-From $EmailFrom "
    If EmailCC <> "" Then _
       Ps1 = Ps1 & "-Cc $EmailCC "
    Ps1 = Ps1 & "-Subject $Subject "
    Ps1 = Ps1 & "-Body $Body "
    Ps1 = Ps1 & "-BodyAsHtml "
    If wAttachedFile <> "" Then _
       Ps1 = Ps1 & "-Attachments $Attachments "
    Ps1 = Ps1 & "-SmtpServer $SMTPServer "
    Ps1 = Ps1 & "-UseSSL "
    Ps1 = Ps1 & "-Credential $Credentials "
    Ps1 = Ps1 & "-Port 587 "

    Call WriteStringToFile("C:\Faas\" & FAUserName & "\Temp\PSEmail.ps1", Ps1)
    Call Shell("powershell -executionpolicy remotesigned -file ""C:\Faas\" & FAUserName & "\Temp\PSEmail.ps1")
    'Powershell's Email Parameters are below
    '[-To] <String[]>
    '[-Subject] <String>
    '[[-Body] <String>]
    '[[-SmtpServer] <String>]
    '[-Attachments <String[]>]
    '[-Bcc <String[]>]
    '[-BodyAsHtml]
    '[-Cc <String[]>]
    '[-Credential <PSCredential>]
    '[-DeliveryNotificationOption <DeliveryNotificationOptions>]
    '[-Encoding <Encoding>]
    '-From <String>
    '[-Port <Int32>]
    '[-Priority <MailPriority>]
    '[-UseSsl]
    '[<CommonParameters>]
End Function
Public Function SendDekraImportEmail(Optional eAttached As String, Optional TestModex As Boolean)
    Dim RST    As DAO.Recordset
    Dim Rst2   As DAO.Recordset
    Dim SQLx, SQLz As String
    Dim m1     As String
    Dim m2     As String
    Dim eSubject As String
    Dim eFromEmail As String
    Dim eToEmail As String
    Dim eCcEmail As String
    Dim eReplyTo As String
    Dim eToLine As String
    Dim iUnitCount, i, Imported, NotImported As Integer
    Dim iUnitSum As Currency
    Dim IncludeBuyersAddr As Boolean
    m1 = ""
    m2 = ""
    eToEmail = "cvanatta@FleetAdvantage.com"
    eCcEmail = "tvanahn@fleetadvantage.com;lmanin@fleetadvantage.com;yjackson@fleetadvantage.com;ljohnson@fleetadvantage.com;afischer@Fleetadvantage.com"
    SQLz = "Select I.*, ClientShortName as Client, UnitDescShortVer as UDesc From InspLinkTemp as I Inner Join vw_SixKeys as S on S.UnitID=I.UnitID"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    Imported = RST.RecordCount
    If Imported = 0 Then GoTo FixMe
    eSubject = "ALERT! Dekra Inspections have been Imported"
    eToLine = "Remarketing Team"
    eFromEmail = "IT@fleetadvantage.com"
    m2 = m2 & "<h3 style=""color:Navy;font-family:Segoe UI;"">List of Imported Inspections:</h3>"
    m2 = m2 & "<b>Please Note:</b> "
    If Imported = 0 Then m2 = m2 & " No Inspections were imported."
    If Imported = 1 Then m2 = m2 & " This " & Imported & " Inspection is "
    If Imported > 1 Then m2 = m2 & " These " & Imported & " Inspections are "
    If Imported > 0 Then m2 = m2 & "set to<b> PENDING</b> and will need to be Reviewed."
    m2 = m2 & "<table class=MsoNormalTable border=0 cellspacing=1 cellpadding=3 border-collapse=collapse >"
    m2 = m2 & "<TR ALIGN=""Center"", bgcolor=""Navy"">"
    m2 = m2 & "<TH>Client</TH>"
    m2 = m2 & "<TH>AssetID</TH>"
    m2 = m2 & "<TH>UnitNum</TH>"
    m2 = m2 & "<TH>VIN</TH>"
    m2 = m2 & "<TH>Ispn Date</TH>"
    m2 = m2 & "<TH>OdoMeter</TH>"
    m2 = m2 & "<TH>Unit Desc</TH>"
    m2 = m2 & "<TH>Cost</TH>"
    m2 = m2 & "<TH>Link</TH>"
    m2 = m2 & "</TR></span>"
    For i = 1 To RST.RecordCount
        If i Mod 2 = 0 Then m2 = m2 & "<TR bgcolor=""WhiteSmoke"">" Else m2 = m2 & "<TR bgcolor=""White"">"
        m2 = m2 & "<TD ALIGN=""Left"">" & RST!Client & "</TD>"
        m2 = m2 & "<TD ALIGN=""Center"">" & RST!UnitID & "</TD>"
        m2 = m2 & "<TD ALIGN=""Left"">" & RST!UnitNum & "</TD>"
        m2 = m2 & "<TD ALIGN=""Left"">" & RST!VIN & "</TD>"
        m2 = m2 & "<TD ALIGN=""Left"">" & RST!InspDate & "</TD>"
        m2 = m2 & "<TD ALIGN=""Left"">" & RST!Odometer & "</TD>"
        m2 = m2 & "<TD ALIGN=""Left"">" & RST!uDesc & "</TD>"
        m2 = m2 & "<TD ALIGN=""Right"">" & Format(RST!iCost, "Currency") & "</TD></span>"
        m2 = m2 & "<TD ALIGN=""Center""><a href=""" & RST!link2 & """>Dekra</a></TD>"
        m2 = m2 & "</span></TR>"
        RST.MoveNext
    Next i
    m2 = m2 & "</table><br /><br />"
    Set Rst2 = CurrentDb.OpenRecordset("Select * From NewDekraInspectionsNoMatch", dbOpenSnapshot)
    On Error Resume Next
    Rst2.MoveLast
    Rst2.MoveFirst
    On Error GoTo 0
    NotImported = Rst2.RecordCount
    If Rst2.RecordCount > 0 Then
        m2 = m2 & "<br /><br /><h3 style=""color:Red;font-family:Segoe UI;"">NOT Imported Inspections:</h3>"
        m2 = m2 & "<b>Please Note:</b>"
        If NotImported = 1 Then m2 = m2 & " This " & NotImported & " Inspection was <b> NOT</b> imported."
        If NotImported > 1 Then m2 = m2 & " These " & NotImported & " Inspections were <b> NOT</b> imported. "
        m2 = m2 & "<table class=MsoNormalTable border=0 bordercolor=black cellspacing=1 cellpadding=3>"
        '  m2 = m2 & "<TR ALIGN=""Center"">"
        m2 = m2 & "<TR ALIGN=""Center"", bgcolor=""Orange"">"
        m2 = m2 & "<TH>AssetID</TH>"
        m2 = m2 & "<TH>UnitNum</TH>"
        m2 = m2 & "<TH>VIN</TH>"
        m2 = m2 & "<TH>Ispn Date</TH>"
        m2 = m2 & "<TH>Unit Desc</TH>"
        m2 = m2 & "<TH>Cost</TH>"
        m2 = m2 & "<TH>Link</TH>"
        m2 = m2 & "</TR></span>"
        For i = 1 To Rst2.RecordCount
            If i Mod 2 = 0 Then m2 = m2 & "<TR bgcolor=""WhiteSmoke"">" Else m2 = m2 & "<TR bgcolor=""White"">"
            m2 = m2 & "<TD ALIGN=""Center"">N/A</TD>"
            m2 = m2 & "<TD ALIGN=""Left"">" & Rst2!iUnitNum & "</TD>"
            m2 = m2 & "<TD ALIGN=""Left"">" & Rst2!VIN & "</TD>"
            m2 = m2 & "<TD ALIGN=""Left"">" & Rst2!iDate & "</TD>"
            m2 = m2 & "<TD ALIGN=""Left"">" & Rst2!Year & " " & Rst2!Make & " " & Rst2!Model & " " & Rst2!Type & "</TD>"
            m2 = m2 & "<TD ALIGN=""Right"">" & Format(Rst2!iCost, "Currency") & "</TD></span>"
            m2 = m2 & "<TD ALIGN=""Center""><a href=""" & Rst2!LinkAddr & """>Dekra</a></TD>"
            m2 = m2 & "</span></TR>"
            Rst2.MoveNext
        Next i
        m2 = m2 & "</table><br /><br />"
    End If
    m1 = m1 & "If you have any questions please contact: <b> IT Dept.</b> email: IT@fleetadvantage.com<br />"

    m1 = m1 & "Please Note: this is an <i>auto-generated</i> email."
    m1 = m1 & "</body></html>"
    If TestModex = True Then     ' set to 1 to test, 0 for normal
        eSubject = "This is a ONLY A TEST! Dekra Inspections have been Imported"
        eToEmail = "ljohnson@fleetadvantage.com"
        eCcEmail = ""
    End If
    m2 = Replace(m2, "'s", "s")
    Call SendServerEmail(eToEmail, eCcEmail, "RMKT " & eSubject, m2, eReplyTo, eAttached)
FixMe:
    On Error Resume Next
    RST.close
    Rst2.close
    Set RST = Nothing
    Set Rst2 = Nothing
    On Error GoTo 0
End Function
Public Function SendBOEmail(eType As String, eBuyersOrder As Long, Optional eAttached As String)
Dim RST As DAO.Recordset
Dim Rst2 As DAO.Recordset
Dim SQLx, SQLz As String
Dim m1 As String
Dim m2 As String
Dim eSubject As String
Dim eFromEmail As String
Dim eToEmail As String
Dim eCcEmail As String
Dim eReplyTo As String
Dim eToLine As String
Dim iUnitCount, i, testmode As Integer
Dim iUnitSum As Currency
Dim IncludeBuyersAddr As Boolean

testmode = 0 ' set to 1 to do testing
m1 = ""
m2 = ""
eCcEmail = "remarketing@fleetadvantage.com;ljohnson@fleetadvantage.com;afischer@Fleetadvantage.com"
'if Fleet Advanatge is the Buyer send TMT the email
If DLookup("BuyerID", "vw_BuyersOrder", "InvoiceRequestID=" & eBuyersOrder & "") = 5453 _
Or DLookup("SalesPerson", "vw_BuyersOrder", "InvoiceRequestID=" & eBuyersOrder & "") = "House Sales" _
   Then eCcEmail = eCcEmail & ";TMT@fleetadvantage.com"


SQLz = "Select * From vw_BuyersOrder Where InvoiceRequestNum=" & eBuyersOrder

SQLx = "SELECT SalesPerson, Email, BuyerShortName, BuyersOrderID As BONum, Count(S.UnitId) As Units, Source "
SQLx = SQLx & " FROM  RMKT_BuyersOrders as BO"
SQLx = SQLx & " INNER JOIN vw_BuyersOrder as S on S.UnitID = BO.UnitID"
SQLx = SQLx & " WHERE BuyersOrderID <> " & eBuyersOrder & " and S.UnitId in"
SQLx = SQLx & "   (SELECT UnitID FROM vw_BuyersOrder as B "
SQLx = SQLx & "    WHERE InvoiceRequestNum=" & eBuyersOrder & ")"
SQLx = SQLx & " GROUP BY SalesPerson,Email, BuyerShortName, BuyersOrderID, Source"


Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
RST.MoveLast
RST.MoveFirst

For i = 1 To RST.RecordCount
   iUnitCount = iUnitCount + 1
   iUnitSum = iUnitSum + RST!SalePrice
   RST.MoveNext
Next i
RST.MoveFirst
'SalesPerson
'Email
'Accepted
'bFullAddress
'DescriptionMed
If eType = "Buyers Order Accepted" Then
   eSubject = "Expect A Payment, Check or Wire Re: Buyers Order# " & RST!InvoiceRequestNum
   eToLine = "Accounting Dept"
   eToEmail = "hprikes@FleetAdvantage.com;ncontreras@FleetAdvantage.com"
   eFromEmail = RST!Email
   eReplyTo = RST!Email
   m2 = "Please review the attached PDF<b> BuyersOrder_" & RST!InvoiceRequestNum & "</b><br/>"
   m2 = m2 & "(<i>This is not the actual buyers order, for reference only</i>)<br/>"
   IncludeBuyersAddr = True
End If

If eType = "Losing BuyersOrders" Then
   eReplyTo = "tvanahn@fleetadvantage.com"
   Set Rst2 = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)

   On Error Resume Next
   Rst2.MoveLast
   Rst2.MoveFirst
   On Error GoTo 0
   If Rst2.RecordCount < 1 Then Exit Function
   RST.MoveLast
   RST.MoveFirst
   eSubject = "ALERT! Units on your Buyers Order(s) have been marked as Pending-Sale (PS)"
   eToLine = "Remarketing Sales"
   eFromEmail = RST!Email
   m2 = m2 & "<h3>List of Pre-Sold Units:</h3>"
   m2 = m2 & "<table class=MsoNormalTable border=1 cellspacing=1 cellpadding=3>"
   m2 = m2 & "<TR ALIGN=""Center"">"
   m2 = m2 & "<TH>AssetID</TH>"
   m2 = m2 & "<TH>VIN</TH>"
   m2 = m2 & "<TH>Description</TH"
   m2 = m2 & "><TH>SoldPrice</TH>"
   m2 = m2 & "<TH>FA Client</TH>"
   m2 = m2 & "</TR></span>"
   For i = 1 To RST.RecordCount
      m2 = m2 & "<TD ALIGN=""Center"">" & RST!UnitID & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & RST!VIN & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & RST!DescriptionMed & "</TD></span>"
      m2 = m2 & "<TD ALIGN=""Right"">" & Format(RST!SalePrice, "Currency") & "</TD>"
      m2 = m2 & "<TD ALIGN=""Center"">" & RST!Source & "</TD>"

      m2 = m2 & "</span></TR>"
      RST.MoveNext
   Next i
   m2 = m2 & "</table><br /><br />"
   m2 = m2 & "<h3>Impacted Buyers Orders:</h3>"
   m2 = m2 & "<table class=MsoNormalTable border=1 cellspacing=1 cellpadding=3>"
   m2 = m2 & "<TR ALIGN=""Center"">"
   m2 = m2 & "<TH>SalesPerson</TH>"
   m2 = m2 & "<TH>Buyer</TH>"
   m2 = m2 & "<TH>BuyersOrder</TH>"
   m2 = m2 & "<TH>Units</TH>"
   m2 = m2 & "<TH>FA Client</TH>"
   m2 = m2 & "</TR></span>"
   For i = 1 To Rst2.RecordCount
      m2 = m2 & "<TD ALIGN=""Left"">" & Rst2!SalesPerson & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & Rst2!BuyerShortName & "</TD>"
      m2 = m2 & "<TD ALIGN=""Center"">" & Rst2!BONum & "</TD></span>"
      m2 = m2 & "<TD ALIGN=""Center"">" & Rst2!UNITS & "</TD>"
      m2 = m2 & "<TD ALIGN=""Center"">" & Rst2!Source & "</TD>"
      m2 = m2 & "</span></TR>"
      eToEmail = eToEmail & Rst2!Email & ";"
      Rst2.MoveNext
    Next i
      m2 = m2 & "</table><br /><br />"
      m2 = m2 & "<b>Please Note:</b> This will<b> invalidate</b> any Buyers Orders with "
      If RST.RecordCount > 1 Then m2 = m2 & "these units on them.<br />" Else m2 = m2 & "this unit on them.<br />"
      m2 = m2 & "Please make the proper adjustments to your Orders, as required.<br />"
      IncludeBuyersAddr = False
End If

If eType = "Units Delivered" Then
   eSubject = "Units Delivered, Re: Buyers Order# " & RST!InvoiceRequestNum
   eToLine = "Remarketing Dept"
   eToEmail = "yjackson@Fleetadvantage.com;afischer@Fleetadvantage.com"
   eFromEmail = RST!Email
   eReplyTo = RST!Email
   m2 = m2 & "<h3>List of Delivered Units:</h3>"
   m2 = m2 & "<table class=MsoNormalTable border=1 cellspacing=1 cellpadding=3>"
   m2 = m2 & "<TR ALIGN=""Center"">"
   m2 = m2 & "<TH>AssetID</TH>"
   m2 = m2 & "<TH>VIN</TH>"
   m2 = m2 & "<TH>Description</TH>"
   m2 = m2 & "<TH>SoldPrice</TH>"
   m2 = m2 & "<TH>FA Client</TH>"
   m2 = m2 & "</TR></span>"
   For i = 1 To RST.RecordCount
      m2 = m2 & "<TD ALIGN=""Center"">" & RST!UnitID & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & RST!VIN & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & RST!DescriptionMed & "</TD></span>"
      m2 = m2 & "<TD ALIGN=""Right"">" & Format(RST!SalePrice, "Currency") & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & RST!Source & "</TD>"
      m2 = m2 & "</span></TR>"
      RST.MoveNext
   Next i
   m2 = m2 & "</table><br /><br />"
   If RST.RecordCount > 1 Then m2 = m2 & "These Units have " Else m2 = m2 & "This Unit has "
   m2 = m2 & "been marked as Sold and <b><u>D</u></b>elivered, Unit Status: ""<b>D</b>"".<br />"
End If
If eType = "BuyersOrderPaid" Then
   eSubject = "Buyers Order Paid Re: Order# " & RST!InvoiceRequestNum
   eToLine = RST!SalesPerson
   eToEmail = RST!Email
   eFromEmail = "hprikes@FleetAdvantage.com"
   eReplyTo = "hprikes@FleetAdvantage.com"
   m2 = m2 & "<br />The Accounting Dept has acknowledged payment for your Buyers Order."
   m2 = m2 & "<h3>List of Paid for Units:</h3>"
   m2 = m2 & "<table class=MsoNormalTable border=1 cellspacing=1 cellpadding=3>"
   m2 = m2 & "<TR ALIGN=""Center"">"
   m2 = m2 & "<TH>AssetID</TH>"
   m2 = m2 & "<TH>VIN</TH>"
   m2 = m2 & "<TH>Description</TH>"
   m2 = m2 & "<TH>Sold Price</TH>"
   m2 = m2 & "<TH>FA Client</TH>"
   m2 = m2 & "</TR></span>"
   For i = 1 To RST.RecordCount
      m2 = m2 & "<TD ALIGN=""Center"">" & RST!UnitID & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & RST!VIN & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & RST!DescriptionMed & "</TD></span>"
      m2 = m2 & "<TD ALIGN=""Right"">" & Format(RST!SalePrice, "Currency") & "</TD>"
      m2 = m2 & "<TD ALIGN=""Left"">" & RST!Source & "</TD>"
      m2 = m2 & "</span></TR>"
      RST.MoveNext
   Next i
   RST.MoveFirst
   m2 = m2 & "</table><br /><br />"
   If RST.RecordCount > 1 Then m2 = m2 & "These Units have " Else m2 = m2 & "This Unit has "
   m2 = m2 & "been marked as <b><u>S</u></b>old, Unit Status: ""<b>S</b>"".<br />"
   m2 = m2 & "<b>Please Note:</b> the<b> Delivery Package</b> has been sent.<br/>Please follow-up with the Buyer: " & RST!BuyerShortName & ", "
   m2 = m2 & "to ensure that the Delivery Receipt has been<b> signed and returned</b> accordingly.<br/>"
End If
'If eType = "" Then eSubject = "Expect A Buyers Order Wire"
RST.MoveFirst
m1 = ""
    m1 = m1 & "<html><body><H3 style=""color:Navy;font-family:Segoe UI;"">" & eSubject & "</H3>"
    m1 = m1 & "To: <b>" & eToLine & "</b><br />"
    m1 = m1 & "Sales Person: <b>" & RST!SalesPerson & "</b><br /><br />"
    m1 = m1 & "Event Date: <b>" & Date & "</b><br />"
    m1 = m1 & "Buyers Order: <b>" & eBuyersOrder & "</b><br />"
    If IncludeBuyersAddr = True Then m1 = m1 & "Buyers Name & Address:<b>" & RST!BuyerAddr & "</b><br />"
    If IncludeBuyersAddr = True Then m1 = m1 & "Buyers Order Unit Count: <b>" & iUnitCount & "</b><br />"
    If IncludeBuyersAddr = True Then m1 = m1 & "Buyers Order Total: <b>" & Format(iUnitSum, "Currency") & "</b><br /><br />"
    m1 = m1 & m2
    If eType = "BuyersOrderPaid" Then
    m1 = m1 & "If you have any questions please contact: <b>Accounting</b> email: " & eFromEmail & "<br />"
    Else
    m1 = m1 & "If you have any questions please contact: <b>" & RST!SalesPerson & "</b> email: " & eFromEmail & "<br />"
    End If
    m1 = m1 & "Please Note: this is an <i>auto-generated</i> email."
    m1 = m1 & "</body></html>"
    If testmode = 1 Then
       eToEmail = "ljohnson@fleetadvantage.com"
       eCcEmail = ""
    End If
    m1 = Replace(m1, "'s", "s")
    Call SendServerEmail(eToEmail, eCcEmail, "RMKT " & eSubject, m1, eReplyTo, eAttached)
   On Error Resume Next
        RST.close
        Rst2.close
FixMe:
    Set RST = Nothing
    Set Rst2 = Nothing
    On Error GoTo 0
End Function
