Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLx, PN, DefaultPrinter, TargetPrinter As String
Dim Printers() As String

Dim n As Long

Private Sub CancelPrint_Click()
DoCmd.OpenReport Me.OpenArgs, acViewReport
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.TargetDocument = Me.OpenArgs
        'Me.PageEnd = Split(Me.OpenArgs, ":")(1)
        Me.PageStart = 1
    End If
    Dim SQLx As String
    DoCmd.SetWarnings False
    SQLx = "Delete * From MyPrinters;"
    DoCmd.RunSQL (SQLx)
    Dim Printers() As String
    Printers = GetPrinterFullNames()
    For n = LBound(Printers) To UBound(Printers)
        If Printers(n) Like "*FA-FILE*" Then
            PN = Split(Printers(n), "FA-FILE\")(1)
            PN = Split(PN, " on")(0)
        Else
            PN = Split(Printers(n), " on")(0)
        End If
        If PN <> "Send To" Then
        SQLx = "INSERT INTO MyPrinters ( PrinterName, PrinterNameLong ) SELECT '" & PN & "' AS PrinterName, '" & Split(Printers(n), " on")(0) & "' As PrinterNameLong;"
        DoCmd.RunSQL (SQLx)
        End If
    Next n
    DoCmd.SetWarnings True
    Me.PrinterList.Requery
    Me.NumOfCopies = 1
End Sub

Private Sub Preview_Click()
If IsNull(Me.PrinterList) Then
MsgBox "Please Select a Printer", vbCritical, "No Printer, No Print"
Exit Sub
End If
DefaultPrinter = Application.Printer.DeviceName
TargetPrinter = Me.PrinterList
Set Application.Printer = Application.Printers(TargetPrinter)
Printer.ColorMode = acPRCMColor
DoCmd.SelectObject acReport, Me.TargetDocument
With Reports(Me.OpenArgs).Printer
If TargetPrinter Like "*Black*" Then .ColorMode = acPRCMMonochrome Else .ColorMode = acPRCMColor
.Copies = Me.NumOfCopies
If Me.PickOrientation = 2 Then .Orientation = acPRORPortrait Else .Orientation = acPRORLandscape
.PrintQuality = acPRPQHigh
DoCmd.OpenReport Me.OpenArgs, acViewPreview
End With
End Sub

Private Sub PrintIt_Click()
If IsNull(Me.PrinterList) Then
MsgBox "Please Select a Printer", vbCritical, "No Printer, No Print"
Exit Sub
End If
DefaultPrinter = Application.Printer.DeviceName
TargetPrinter = Me.PrinterList
Set Application.Printer = Application.Printers(TargetPrinter)
Printer.ColorMode = acPRCMColor
DoCmd.SelectObject acReport, Me.TargetDocument
With Reports(Me.OpenArgs).Printer
If TargetPrinter Like "*Black*" Then .ColorMode = acPRCMMonochrome Else .ColorMode = acPRCMColor
.Copies = Me.NumOfCopies
If Me.PickOrientation = 2 Then .Orientation = acPRORPortrait Else .Orientation = acPRORLandscape
.PrintQuality = acPRPQHigh
End With
DoCmd.PrintOut acSelection, Me.PageStart, Me.PageEnd
Set Application.Printer = Application.Printers(DefaultPrinter)
DoCmd.close acForm, Me.Form.Name
End Sub
