Attribute VB_Name = "SharePoint"
Option Compare Database
Option Explicit

Public Function GetSchedulDirectoryOLd(SearchString As String) As String

            Dim wb As WebDriver
            Dim DirURL As String
            Dim Committ As String
           Set wb = New Selenium.EdgeDriver
           ' Set WB = New Selenium.PhantomJSDriver
            wb.start
            wb.Get "https://fleetadvantage.sharepoint.com/Customers/Forms/AllItems.aspx"
            'WB.Wait (2000)
            wb.FindElementByXPath("/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div/div[2]/div/div[2]/div/div/form/input").SendKeys SearchString
            wb.FindElementByXPath("/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div/div[2]/div/div[2]/div/div/form/span[6]/button").Click
On Error GoTo FixMe
            wb.FindElementByXPath("/html/body/div[1]/div/div[2]/div[1]/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div[1]/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[1]/div/div/div[2]/div[2]/div/div[1]/span/span/button").Click
              Committ = MsgBox("Is the OPEN Webage the directory you want?: " & _
                        vbNewLine & vbNewLine & "If so click 'YES'", vbInformation + vbYesNo, "Look before you Jump")
             If Committ = vbNo Then DirURL = "" Else DirURL = wb.URL()
            wb.close
           GetSchedulDirectory = DirURL
FixMe:
On Error GoTo 0
End Function
Public Function GetSchedulDirectory(SearchString As String) As String
    Dim FlowPath As String
    Dim ResultPath As String
    Dim Result As String

    ' Dynamic per-user path
    FlowPath = Environ("LOCALAPPDATA") & "\Microsoft\Power Automate Desktop\Flow\{GUID}\flow.msf"

    ' Dynamic return file path
    ResultPath = "C:\Faas\" & ("UserName") & "\Temp\PAD_Return.txt"

    ' Delete old result file if it exists
    If Dir(ResultPath) <> "" Then Kill ResultPath

    ' Launch PAD flow with input + output path
    Shell """" & FlowPath & """ --input SearchString=""" & SearchString & """ --input ResultPath=""" & ResultPath & """", vbNormalFocus

    ' Wait for PAD to finish (simple loop)
    Do While Dir(ResultPath) = ""
        DoEvents
    Loop

    ' Read the result
    Result = ReadAllText(ResultPath)

    GetSchedulDirectory = Result
End Function
Public Function ReadAllText(FilePath As String) As String
    Dim f As Integer
    f = FreeFile
    Open FilePath For Input As #f
    ReadAllText = Input$(LOF(f), f)
    Close #f
End Function
Public Function GetSharePointLink(URL As String, Optional JustDir As Boolean, Optional isTest As Boolean) As String
    Dim RawLink As String
    Dim EndLink As String
    Dim DocTypeX As String
    Dim WebPrefix As String
    Dim Filex As String
    Dim ParentX As String
    Dim ie As Variant
    Dim WebString As String
    Dim Tempx As String
    Dim Linktype As String
    Dim i As Integer
    Dim SufixLen As Integer

    RawLink = URL
    If URL Like "*AllItems.aspx*" Then Linktype = "AllType"
    If URL Like "*/g/*" Then Linktype = "AllType"
        If URL Like "*id=%2F*" Then 'And Linktype = "" Then
            Tempx = SplitString(URL, "id=%2F", 1)
            Tempx = Replace(Tempx, "%2F", "\")
            Tempx = Replace(Tempx, "%20", " ")
            Tempx = Replace(Tempx, "%2D", "-")
            Tempx = Replace(Tempx, "%28", "(")
            Tempx = Replace(Tempx, "%29", ")")
            Tempx = Replace(Tempx, "%5F", "_")
            Tempx = Replace(Tempx, "%2E", ".")
            Tempx = Replace(Tempx, "%26", "&")
            If Tempx Like "*parent=\*" Then Tempx = SplitString(Tempx, "parent=\", 1)
            If Tempx Like "*&viewid*" Then Tempx = SplitString(Tempx, "&viewid", 0)
            If Right(ParentX, 1) = "#" Then ParentX = Left(ParentX, Len(ParentX) - 1)
            ParentX = SplitString(ParentX, "#", 0) & "\"
            Filex = SplitString(URL, "&parent=", 0)
            Filex = SplitString(Filex, ParentX, 1)
            'DocTypeX = SplitString(Filex, ".", 1)
            If Left(Right(Filex, 3), 1) = "." Then DocTypeX = UCase(Right(Filex, 3)) Else DocTypeX = "ukn"
            If ParentX Like "*Error*" Then
                Linktype = "Directory"
                If Tempx Like "*Documentation\*" Then Filex = SplitString(Tempx, "Documentation\", 1) Else Filex = Replace(Tempx, "Customers\", "")
                ParentX = Tempx
                DocTypeX = "DIR"
                'URL = Filex & "#" & URL & "#"
                GoTo LinkComplete
            End If
        Else
            Dim wb As WebDriver
            Set wb = New Selenium.EdgeDriver
            ' Set WB = New Selenium.PhantomJSDriver
            wb.start
            If Linktype = "AllType" Then wb.Get RawLink Else wb.Get URL
            WebString = wb.PageSource
            If isTest = True Then GoTo WriteMe
            If Linktype = "AllType" Then
                If WebString Like "*RootFolder*" Then
                    Tempx = SplitString(WebString, "&RootFolder=%2F", 1)
                    Tempx = SplitString(Tempx, "&", 0)
                    Tempx = Replace(Tempx, "%2F", "\")
                    Tempx = Replace(Tempx, "%20", " ")
                    Tempx = Replace(Tempx, "%2D", "-")
                    Tempx = Replace(Tempx, "%28", "(")
                    Tempx = Replace(Tempx, "%29", ")")
                    Tempx = Replace(Tempx, "%5F", "_")
                    Tempx = Replace(Tempx, "%2E", ".")
                    Tempx = Replace(Tempx, "%26", "&")

                    If Linktype = "AllType" Then
                        i = HowManyInStr(Tempx, "\")
                        Filex = Split(Tempx, "\")(i)
                        'DocTypeX = UCase(SplitString(Filex, ".", 1))
                        If Left(Right(Filex, 5), 1) = "." Then DocTypeX = UCase(Right(Filex, 4)) Else DocTypeX = "ukn"
                        If DocTypeX = "ukn" And Left(Right(Filex, 4), 1) = "." Then DocTypeX = UCase(Right(Filex, 3)) Else DocTypeX = "ukn"

                    End If
                    ParentX = Tempx
                Else
                    Filex = SplitString(SplitString(wb.FindElementByCss("title").Attribute("outerHTML"), "</title>", 0), "<title>", 1)
                    SufixLen = 0
                    If Left(Right(Filex, 5), 1) = "." Then SufixLen = 4
                    If Left(Right(Filex, 4), 1) = "." Then SufixLen = 3
                    If SufixLen > 0 Then DocTypeX = UCase(Right(Filex, SufixLen)) Else DocTypeX = "ukn"
                End If
                wb.close
                If DocTypeX = "ukn" Or DocTypeX Like "*Error:*" Then
                    Linktype = "Directory"
                    If Tempx Like "*Documentation\*" Then
                        Filex = SplitString(Tempx, "Documentation\", 1)
                    Else
                        Filex = Replace(Tempx, "Customers\", "")
                    End If
                    ParentX = Tempx
                    DocTypeX = "DIR"
                    URL = "https://fleetadvantage.sharepoint.com/" & Tempx
                    GoTo LinkComplete
                End If
                'If WebString Like "*appName = 'Word';*" Or WebString Like "*var appName = 'Excel';*" Then
                If WebString Like "*appName =*" Then
                    Filex = SplitString(SplitString(WebString, "</title>", 0), "<title>", 1)
                    If DocTypeX = "ukn" And Left(Right(Filex, 5), 1) = "." Then DocTypeX = UCase(Right(Filex, 4))
                    If DocTypeX = "ukn" And Left(Right(Filex, 4), 1) = "." Then DocTypeX = UCase(Right(Filex, 3))
                    ParentX = SplitString(WebString, "ParentFolderFullUrl"":", 1)
                    ParentX = SplitString(ParentX, "DocAgeBucketAtViewTimeBasedOnLastModifiedTime", 0)
                    ParentX = Replace(ParentX, """", "")
                    ParentX = Replace(ParentX, ",", "")
                    ParentX = Replace(ParentX, "https://fleetadvantage.sharepoint.com/", "")
                   ' URL = "https://fleetadvantage.sharepoint.com/" & Replace(Filex, "\", "/")
                Else
                   If Not URL Like "*https://fleetadvantage.sharepoint.com/*" Then URL = "https://fleetadvantage.sharepoint.com/" & Replace(Tempx, "\", "/")
                End If

            End If
            If Linktype <> "AllType" Then
                Filex = SplitString(WebString, "<input name=""fileName""", 1)
                Filex = SplitString(Filex, "<input name=""fileSize""", 0)
                Filex = Replace(Filex, """", "'")
                Filex = SplitString(Filex, "value='", 1)
                Filex = SplitString(Filex, "'>", 0)
                'DocTypeX = SplitString(Filex, ".", 1)
                If Left(Right(Filex, 4), 1) = "." Then DocTypeX = UCase(Right(Filex, 3)) Else DocTypeX = "ukn"
                ParentX = SplitString(WebString, "ParentFolderFullUrl"":", 1)
                ParentX = SplitString(ParentX, "DocAgeBucketAtViewTimeBasedOnLastModifiedTime", 0)
                ParentX = Replace(ParentX, """", "")
                ParentX = Replace(ParentX, ",", "")
                ParentX = Replace(ParentX, "https://fleetadvantage.sharepoint.com/", "")
            End If
    End If
LinkComplete:
    GetSharePointLink = "|" & ParentX & "|" & Filex & "|" & DocTypeX & "|" & URL

    Exit Function
WriteMe:
    Dim intFileNum As Integer
    intFileNum = FreeFile
    'change Output to Append if you want to add to an existing file
    ' rather than creating a new file each time
    Open "C:\Faas\" & FAUserName & "\Temp\WebText.txt" For Output As intFileNum
    Print #intFileNum, WebString
    Close intFileNum

End Function
