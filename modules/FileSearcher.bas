Attribute VB_Name = "FileSearcher"
Option Compare Database
Option Explicit
Sub KillFiles(sPath As String, sFile As String)
'**** This will delete / kill files that match the sPath and sFiles name strings
    Dim iFilesNum As Integer
    Dim iCount As Integer
    Dim sPathArry() As String
    Dim iFilesFound As Boolean
    Call MessageUser("Open", "searching...", "working...")
    iFilesFound = FindFiles(sPath, sPathArry, iFilesNum, sFile, True)
    If iFilesFound Then
        For iCount = 1 To iFilesNum
            Call MessageUser("Update", "Deleting Targeted Files:", sPathArry(2, iCount))
            Kill (sPathArry(1, iCount) & sPathArry(2, iCount))
        Next
    End If
    Call MessageUser("Close")
End Sub

Sub testSearchFiles()
    Dim iFilesNum As Integer
    Dim iCount As Integer
    Dim sPathArry() As String
    Dim iFilesFound As Boolean
    Dim sPath As String
    Dim sSearch As String
    sPath = "C:\Users\" & FAUserName & "\Downloads"
    sSearch = "Schedule_" & DatePart("YYYY", Date) & "-"
    If DatePart("M", Date) < 10 Then sSearch = sSearch & "0" & DatePart("M", Date) & "-" Else sSearch = sSearch & DatePart("M", Date) & "-"
    If DatePart("D", Date) < 10 Then sSearch = sSearch & "0" & DatePart("D", Date) Else sSearch = sSearch & DatePart("D", Date)
    sSearch = sSearch & "_*.xlsx"
    iFilesFound = FindFiles(sPath, sPathArry, iFilesNum, sSearch, True)
    If iFilesFound Then
        For iCount = 1 To iFilesNum
            MsgBox "Path: " & sPathArry(1, iCount) & vbNewLine & _
                   vbNewLine & "File name: " & sPathArry(2, iCount), _
                   vbInformation, "Files Found"
        Next
    End If
End Sub
Function FindFiles(ByVal sPath As String, _
    ByRef sFoundFiles() As String, _
    ByRef iFilesFound As Integer, _
    Optional ByVal sFileSpec As String = "*.*", _
    Optional ByVal blIncludeSubFolders As Boolean = False) As Boolean
' FindFiles
' ---------
' Find all files matching the specified file spec, starting from the specified path
' and search subfolder as required.
'
' Parameters
' ----------
' sPath (String): Startup folder, e.g. "C:\Users\Username\Documents"
'
' sFoundFiles (String): Two dimensional array to store the path and name of found files.
'   The dimension of this array is (1 To 2, 1 To nnn), where nnn is the number of found
'   files. The elements of the array are:
'      sFoundFiles(1, xxx) = File path     (xxx = 1 to nnn)
'      sFoundFiles(2, xxx) = File name
'
' iFilesFound (Integer): Number of files found.
'
' sFileSpec (String): Optional parameter with default value = "*.*"
'
' blIncludeSubFolders (Boolean): Optional parameter with default value = False
'   (which means sub-folders will not be searched)
'
' Return values
' -------------
' True: One or more files found, therefore
'   sFoundFiles = Array of paths and names of all found files
'   iFilesFound = Number of found files
' False: No files found, therefore
'   iFilesFound = 0
'
' **********************************  Important Note  **********************************
'
' When searching for *.xls, FindFiles returns, in addition to xls files, xls* (not xls?)
' files (e.g. xlsX, xlsM, xlsWhatever, etc.). The reason is that FindFiles uses the Dir
' function and these files are returned by Dir! The most interesting thing here is that
' Windows search (including DOS DIR command) returns the same! It seems Excel Dir uses
' Windows search without any further checking or refinements.
'
' This is also true for *.doc and *.ppt files. Actually, this is true whenever a
' three-character file extension is specified; *.txt, *.pdf, *.x?s, etc.
'
' Moreover, if the last character of the specified extension was a question mark (?) or
' an asterisk (*), the returned files would be the same (e.g. *.txt? and *.txt* return
' the same files). This means, files with more than four-character extension are returned
' in both cases. This is exactly the same behaviour when specifying three-character
' extension (*.txt)…so weird!
'
' The aforementioned behaviour was observed in Windows 7 using Excel 2010 (mostly, Excel
' is not a key player here).
'
' Not everything is covered in this note as further tests might reveal more. So, keep
' these things in mind when using Dir or FindFile.
'
' Constructive comments and Reporting of bugs would be appreciated.
'
' **************************************************************************************
'
' Using the function (sample code)
' --------------------------------
' Dim iFilesNum As Integer
' Dim iCount As Integer
' Dim sPathArry() As String
' Dim iFilesFound As Boolean
'
' iFilesFound = FindFiles("C:\Users\Username\Documents", _
'     sPathArry, iFilesNum, "*.xls", True)
' If iFilesFound Then
'     For iCount = 1 To iFilesNum
'         MsgBox "Path: " & sPathArry(1, iCount) & vbNewLine & _
'             vbNewLine & "File name: " & sPathArry(2, iCount), _
'             vbInformation, "Files Found"
'     Next
' End If

    Dim iCount As Integer           '* Multipurpose counter
    Dim sFileName As String         '* Found file name
    '* FileSystem objects
    Dim oFileSystem As Object, _
        oParentFolder As Object, _
        oFolder As Object

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set oParentFolder = oFileSystem.GetFolder(sPath)
    If oParentFolder Is Nothing Then
        FindFiles = False
        On Error GoTo 0
        Set oParentFolder = Nothing
        Set oFileSystem = Nothing
        Exit Function
    End If
    sPath = IIf(Right(sPath, 1) = "\", sPath, sPath & "\")
    '*
    '* Find files[/COLOR]
    sFileName = Dir(sPath & sFileSpec, vbNormal)
    Do While sFileName <> ""
        iCount = UBound(sFoundFiles, 2)
        iCount = iCount + 1
        ReDim Preserve sFoundFiles(1 To 2, 1 To iCount)
        sFoundFiles(1, iCount) = sPath
        sFoundFiles(2, iCount) = sFileName
        sFileName = Dir()
    Loop
    If blIncludeSubFolders Then
        '*
        '* Select next subforbers[/COLOR]
        For Each oFolder In oParentFolder.SubFolders
            FindFiles oFolder.path, sFoundFiles, iFilesFound, sFileSpec, blIncludeSubFolders
        Next
    End If
    FindFiles = UBound(sFoundFiles, 2) > 0
    iFilesFound = UBound(sFoundFiles, 2)
    On Error GoTo 0
    '*
    '* Clean-up[/COLOR]
    Set oFolder = Nothing
    Set oParentFolder = Nothing
    Set oFileSystem = Nothing

End Function
