Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(Cancel As Integer)
    If Not IsNull(Me.OpenArgs) Then Me.Report.RecordSource = Me.OpenArgs
End Sub
