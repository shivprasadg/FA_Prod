Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(Cancel As Integer)
Me.RecordSource = Nz(Me.OpenArgs, "")
If Me.OpenArgs Like "*GroupID=*" Then Me.ReportHeader.caption = "Final Syndication Economics (single)" Else Me.ReportHeader.caption = "Final Syndication Economics"
End Sub
