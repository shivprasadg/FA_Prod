Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub PrintAll3_Click()
DoCmd.OpenReport "RmktBOSRpt"
DoCmd.OpenReport "RmktTitleRelLetter"
DoCmd.OpenReport "RmktTitleCoverLetter"
DoCmd.close acReport, "RmktTitleCoverLetter"
DoCmd.close acReport, "RmktTitleRelLetter"
DoCmd.close acReport, Me.Report.Name
End Sub
