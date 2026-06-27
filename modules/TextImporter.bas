Attribute VB_Name = "TextImporter"
Option Compare Database
Option Explicit

Sub LoadText()
Dim fso As Scripting.FileSystemObject
Dim f As File
Dim FilePath, rpttext As String
Dim F2 As TextStream
Dim ts As TextStream
Dim lngCount, Linei As Long
Dim GetNext As Integer
Dim PlateVin As String
Dim RecordNum As Integer


 Set fso = New Scripting.FileSystemObject
'FilePath = "C:\Users\ljohnson\Desktop\PDFIn\testme.txt"
' intFileNum = FreeFile
' Open FilePath For Input As intFileNum
FilePath = "C:\Users\ljohnson\Desktop\PDFIn\Walmart_Reg.txt"
Set f = fso.GetFile(FilePath)
Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)
PlateVin = ""
'ts = ""
GetNext = 0
Do
rpttext = ts.ReadLine
Linei = Linei + 1
If GetNext = 1 Then
RecordNum = RecordNum + 1
PlateVin = PlateVin & RecordNum & ","
PlateVin = PlateVin & Trim(SplitString(rpttext, " ", 0)) & ", "
PlateVin = PlateVin & Trim(SplitString(SplitString(rpttext, " ", 1), " ", 0)) & ", "
GetNext = 0
End If

If rpttext = "Plate  VIN  Unit  Year  Make  Type  Axles  Seats  Fuel" Then
GetNext = 1
End If
'If rpttext = "Whatever criteria, or even if instr(rpttext,'whateveryouwant') " > 0 Then
   ' do something with rpttext
Loop
 Close f
End Sub
Sub ReadFile()

'again, we need this strange thing to exist so that ...

Dim fso As New FileSystemObject

'the file we're going to read from

Dim ts As TextStream

'... we can open a text file with reference to it

Set ts = fso.OpenTextFile("C:\Wise Owl\info.txt", ForReading)

'keep reading in lines till no more

Dim ThisLine As String

Dim i As Integer

i = 0

Do Until ts.AtEndOfStream

ThisLine = ts.ReadLine

i = i + 1

Debug.Print "Line " & i, ThisLine

Loop

'close down the file

ts.close

End Sub
Sub TextStreamTest()
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateTextFile "test1.txt"            'Create a file
    Set f = fs.GetFile("test1.txt")
    Set ts = f.OpenAsTextStream(ForWriting, TristateUseDefault)
    ts.Write "Hello World"
    ts.close
    Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)
    s = ts.ReadLine
    MsgBox s
    ts.close
End Sub
