Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Public ReportX, Queryx As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Close()
    DoCmd.OpenForm " FAMainMenu"
End Sub

Private Sub Form_Load()
    Me.Search_Client = ""
    Me.Search_Client.SetFocus
    Call Search_Client_Change
    Me.UnitsList.Visible = False
    Me.UnitListcount.Visible = False
    Me.line2.Visible = False
    Me.SUFbyUnit.Visible = False
    Me.SearchVin.Visible = False
    ReportX = "SUFReport"
    Queryx = "vw_SUFReportQuery"
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub

Public Sub Company_List_Click()
    Me.LocationList.SetFocus
    Me.LocationList = ""
    Me.LocationList.Requery
    Me.ScheduleList.SetFocus
    Me.ScheduleList = ""
    Me.ScheduleList.Requery
    Me.GroupList = ""
    Me.GroupList.Requery
    Me.UnitsList.Visible = False
    Me.UnitListcount.Visible = False
    Me.line2.Visible = False
    Me.SearchVin.Visible = False
    Me.SearchVin = ""
    Me.SUFbyUnit.Visible = False
    Call ClearList(Me.LocationList.Name)
    Call ClearList(Me.ScheduleList.Name)
End Sub
Public Sub LocationList_Click()
    Call ClearList(Me.ScheduleList.Name)
    Me.UnitsList.Visible = True
    Me.UnitListcount.Visible = True
    Me.line2.Visible = True
    Me.SearchVin.Visible = True
    Me.SUFbyUnit.Visible = True
    Me.UnitsList = ""
    Me.SearchVin = ""

    ReportX = "SUFReport"

    SQLz = ""
    SQLz = SQLz & "SELECT clientID, UnitDesc, unitvin AS VIN, UnitID, InspDate, ParkDate "
    SQLz = SQLz & "FROM vw_SUFReportQuery "
    SQLz = SQLz & "WHERE LocationID=" & [LocationList] & " "
    SQLz = SQLz & "GROUP BY clientID, UnitDesc, unitvin, UnitID, InspDate, ParkDate "
    SQLz = SQLz & "ORDER BY unitvin;"
    Me.UnitsList.RowSource = SQLz
End Sub

Private Sub ScheduleList_Click()
    Call ClearList(Me.LocationList.Name)
    Me.UnitsList.Visible = True
    Me.UnitListcount.Visible = True
    Me.line2.Visible = True
    Me.SearchVin.Visible = True
    Me.SUFbyUnit.Visible = True
    Me.UnitsList = ""
    Me.SearchVin = ""
    Me.GroupList = ""
    Me.GroupList.Requery

    ReportX = "SUFReport"

    SQLz = ""
    SQLz = SQLz & "SELECT clientID, UnitDesc, unitvin AS VIN, UnitID, InspDate, ParkDate "
    SQLz = SQLz & "FROM vw_SUFReportQuery "
    SQLz = SQLz & "WHERE SchID=" & Me.ScheduleList & " "
    SQLz = SQLz & "GROUP BY clientID, UnitDesc, unitvin, UnitID, InspDate, ParkDate "
    SQLz = SQLz & "ORDER BY unitvin;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub GroupList_Click()

        ReportX = "SUFReport"
    SQLz = ""
    SQLz = SQLz & "SELECT clientID, UnitDesc, unitvin AS VIN, UnitID, InspDate, ParkDate "
    SQLz = SQLz & "FROM vw_SUFReportQuery "
    SQLz = SQLz & "WHERE SgrpId=" & Me.GroupList & " "
    SQLz = SQLz & "GROUP BY clientID, UnitDesc, unitvin, UnitID, InspDate, ParkDate "
    SQLz = SQLz & "ORDER BY unitvin;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub Search_Client_Change()
    SQLz = "SELECT clientID, ClientShortName, Count(unitvin) AS Units " & _
           "From vw_SUFReportQuery GROUP BY clientID, ClientShortName " & _
           "HAVING ClientShortName Like '*" & Me.Search_Client.Text & "*' " & _
           "Order By ClientShortName;"
    Me.Company_List.RowSource = SQLz
End Sub

Private Sub SearchVIN_Change()
    SQLz = "SELECT clientID, UnitDesc, unitvin AS VIN, InspDate AS Inspected, ParkDate AS Parked " & _
           "FROM vw_SUFReportQuery " & _
           "WHERE LocationID=" & [LocationList] & " AND unitvin Like '*" & Me.SearchVin.Text & "*' " & _
           "GROUP BY vw_SUFReportQuery.clientID, vw_SUFReportQuery.UnitDesc, vw_SUFReportQuery.unitvin, vw_SUFReportQuery.InspDate, vw_SUFReportQuery.ParkDate " & _
           "ORDER BY unitvin;"
    Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SUFbyGroup_Click()
    Me.LocationList = 0
    If Me.GroupList & "test" = "test" Then
        MsgBox "Please Select a Group", vbInformation, "Nothing = Nothing"
        Exit Sub
    End If
    Call MakeSUF("Group")
End Sub

Private Sub SUFbyUnit_Click()
    If Me.UnitsList & "test" = "test" Then
        MsgBox "Please Select a Unit", vbInformation, "Nothing = Nothing"
        Exit Sub
    End If
    Call MakeSUF("Unit")
End Sub
Private Sub SUFbyLocation_Click()
    If Me.LocationList & "test" = "test" Then
        MsgBox "Please Select a Location", vbInformation, "Location, Location, Location"
        Exit Sub
    End If
    Call MakeSUF("Location")
End Sub

Private Sub SUFbySchedule_Click()
    Me.LocationList = 0
    If Me.ScheduleList & "test" = "test" Then
        MsgBox "Please Select a Schedule", vbInformation, "Nothing = Nothing"
        Exit Sub
    End If
    Call MakeSUF("Sch")
End Sub

Public Sub MakeSUF(ReportType As String)
    Call CloseReports
    If Me.Company_List.Column(1) Like "*mclane*" Or Me.Company_List.Column(1) Like "*Transco*" Then ReportX = "SUFReport_Mclane"
    If Me.Company_List.Column(1) Like "*maines*" Then ReportX = "SUFReport_Maines"

    If DCount("ClientID", "SUFConditions", "ClientID=" & Me.Company_List.Column(0)) > 0 Then ReportX = "SUFwConditions"
    SQLz = "SELECT * FROM vw_SUFReportQuery "
    If ReportType = "Unit" Then SQLz = SQLz & " WHERE UnitID = " & Me.UnitsList.Column(3) & " "
    If ReportType = "Location" Then SQLz = SQLz & " WHERE LocationID = " & Me.LocationList & " "
    If ReportType = "Sch" Then SQLz = SQLz & " WHERE SchID = " & Me.ScheduleList & " "
    If ReportType = "Group" Then SQLz = SQLz & "WHERE sGrpID = " & Me.GroupList & " "
    SQLz = SQLz & " Order By UnitUnitNum;"
    DoCmd.OpenReport ReportX, acViewReport, , , , SQLz
End Sub
Private Sub CloseReports()
    DoCmd.close acReport, "SUFReport"
    DoCmd.close acReport, "SUFReport_Mclane"
    DoCmd.close acReport, "SUFReport_Maines"
    DoCmd.close acReport, "SUFReport_AutoHauler"
End Sub
Public Sub ClearList(ListBoxName As String)
    Dim ItemsX As Variant
    With Me(ListBoxName)
        For Each ItemsX In .ItemsSelected
            .Selected(ItemsX) = False
        Next ItemsX
        Me(ListBoxName) = Null
    End With
End Sub

Private Sub UnitsList_DblClick(Cancel As Integer)
    If Me.UnitsList.Column(0) & "test" = "test" Then
        MsgBox "There must select a unit before you can view it", , "Really?"
        Exit Sub
    End If
    Call OpenUnitsForm(Me.UnitsList.Column(3))
    DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitsList.Column(3) & ""
    Me.SearchVin.SetFocus
End Sub
