Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public SQLz, SQLx As String

Private Sub EXIT_Click()
DoCmd.close acReport, Me.Report.Name
End Sub

Private Sub ExportMe_Click()
    Dim XLS, Headerx As String
    XLS = ""
    SendReport2Excel SQLz, "OutstandingATF", "Outstanding ATF " & Date, True, XLS
End Sub
Private Sub Report_Load()
    SQLz = "Select * From OutstandingATFSchedule_KPI Where GroupAdded ='Pending'"
    Me.Report.RecordSource = SQLz
End Sub

Private Sub Printme_Click()
DoCmd.close acForm, "Printers"
DoCmd.OpenForm "Printers", acNormal, , , , , Me.Report.Name
Forms!Printers.PageEnd = Me.Report.Pages
Forms!Printers.PickOrientation = 1
End Sub
