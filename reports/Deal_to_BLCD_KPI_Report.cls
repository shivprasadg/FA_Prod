Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub EXIT_Click()
DoCmd.close acReport, Me.Report.Name
End Sub

Private Sub ExportMe_Click()
    Dim XLS, Headerx As String
    'HeaderX = "EXIT Report - " & Me.ReportTitle.Caption & " - " & Format(Date, "mm-dd-yy")
    'If Me.OpenArgs Like "*Month*" Then HeaderX = HeaderX & " - " & Me.MonthsUnitEOL
    'If Me.OpenArgs Like "*Location*" Then HeaderX = HeaderX & " - " & Me.LocationName
    'If Me.OpenArgs Like "*Client*" Then HeaderX = HeaderX    '& " - " & Me.CompanyName
    'If Me.OpenArgs Like "*BDEX*" Then HeaderX = HeaderX & " - " & Me.SalespersonName
    '_________________________
    'XLS = Col letter,Format Option
    ' first set = number of columns
    'Format Options Type Command
    'Types:     T = Text | B = Bold | $ = Currency
    'Commands:  C = Center | L=Lexf | R=Right | E = Sum col | T= Total Col Count | X = COUNT

    XLS = ""
    SendReport2Excel "Select * From Deal_KPI_Report", "Deal_KPI_Report", "Deal To BLCD Report " & Date, True, XLS
End Sub



Private Sub Printme_Click()
DoCmd.PrintOut
End Sub
