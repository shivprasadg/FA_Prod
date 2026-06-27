Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ArchiveReport_Click()
    Dim pFileName As String
    Dim sFileName As String
    Dim tt     As Integer
    pFileName = "P:\FLEETADVANTAGE\Economics\ATLAAS TFR\" & DatePart("yyyy", Date) & "\TFRDetail_" & Format(Date, "mm_dd_yyyy") & ".pdf"
    If AskToDoTask("ATLAAS will Archive a copy of the TFR" & vbNewLine & "as: " & SplitString(pFileName, DatePart("yyyy", Date) & "\", 1) & vbNewLine) = True Then
CheckFile:
        tt = tt + 1
        If FileExists(pFileName) = True Then
             If tt = 1 Then pFileName = SplitString(pFileName, ".pdf", 0) & "(1).pdf"
             If tt > 1 Then pFileName = SplitString(pFileName, "(", 0) & "(" & tt & ").pdf"
            GoTo CheckFile
        End If

        DoCmd.OutputTo acOutputReport, "TFRCompare", acFormatPDF, pFileName
    End If
End Sub

Private Sub Sch_Grp_DblClick(Cancel As Integer)
DoCmd.close acForm, "GroupForm"
DoCmd.OpenForm "GroupForm", , , "SGrpID = " & Me.SGrpID
End Sub
