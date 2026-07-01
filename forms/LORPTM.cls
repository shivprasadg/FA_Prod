Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Dim XLS As String
Dim Headerx As String


Private Sub AllUnitsLeased_Click()
'DoCmd.OpenQuery "AllLeasedUnits_byLeaseDate"
AddReportLog "AllUnitsLeased_Click"
If Me.EndDateRange & "test" = "test" Or Me.StartDateRange & "test" = "test" Then
MsgBox "You must select a Start and End Date", vbInformation, "Slow Down..."
Exit Sub
End If
    Dim ReportX As String
    Dim XLS As String
    Dim Countx As Integer
 If Me.AllUnitsDateType = "Accepted Date" Then ReportX = "AllLeasedUnits_byAcceptedDate"
 If Me.AllUnitsDateType = "BLCD Date" Then ReportX = "AllLeasedUnits_byLeaseDate"
 If Me.AllUnitsDateType = "Invoiced Date" Then ReportX = "AllLeasedUnits_byInvoiceDate"
    Countx = DCount("UnitId", "" & [ReportX] & "")
    XLS = "|B,BC|C,BC|E,TL|I,BC|J,TC|K,IC|J,TX|O,BC|P,TX|AB,$E|AC,$E|AD,$E|"
 AddReportLog "AllUnitsLeased_Click"
 FormatExcel ReportX, "C:\Faas\" & FaUserName & "\Temp\" & ReportX, "LeasedUnits", , XLS, Countx, "Leased Units by " & AllUnitsDateType, True
End Sub

Private Sub ClearAll_Click()
    Me.StartDateRange = Null
    Me.EndDateRange = Null
    Me.SelectClient = Null
    Me.AllUnitsDateType = Null
End Sub

Private Sub CloseMe_Click()
DoCmd.close acForm, Me.Form.Name
AddReportLog "CloseMe_Click"
End Sub

Private Sub OrderedOwnedSyn_Click()
    Dim XLS As String
    Dim Countx As Integer
    SQLz = "Select * From OrderedOwnedSyn_LeaseReport"
    MessageUser "Open", "Getting Ord-Owned-Syn Report", "This will take a second or two"
    XLS = "|C,BL|D,BC|F,BL|J,XE|K,$E|L,$E|M,BL|N,XE|O,$E|P,$E|Q,BL|R,XE|S,$E|T,$E|"
    AddReportLog "OrderedOwnedSyn_Click"
    Call SendReport2Excel(SQLz, "Ord-Owned-Syn", "Ord-Owned-Syn: (" & Date & ")", True, XLS, 0, True)
    MessageUser "Close"
  End Sub

Private Sub MiniExitReport_Click()
    Headerx = "Mini Exit Report: (" & Date & ")"

    SQLz = "Select Client "
    SQLz = SQLz & ",Sch"                        'Col B
    SQLz = SQLz & ",[Group] as GRP"
    SQLz = SQLz & ",AssetID"
    SQLz = SQLz & ",UnitNum"
    SQLz = SQLz & ",UnitVIN as [Unit VIN]"       'Col G
    SQLz = SQLz & ",[Status]"
    SQLz = SQLz & ",InService AS [In Service]"
    SQLz = SQLz & ",Garage"                      'Col I
    SQLz = SQLz & ",State"
    SQLz = SQLz & ",UnitYr as MYear"
    SQLz = SQLz & ",UnitMake as Make"
    SQLz = SQLz & ",UnitModel as Model"
    SQLz = SQLz & ",UnitType as UType"           'Col N
    SQLz = SQLz & ",UnitSubType as SubType"
    SQLz = SQLz & ",Axle"
    SQLz = SQLz & ",AxleConfig"                  'Col Q
    SQLz = SQLz & ",UnitConfig"
    SQLz = SQLz & ",Engine"
    SQLz = SQLz & ",Trans"                       'Col T
    SQLz = SQLz & ",OEC AS [Total OEC Cost]"
    SQLz = SQLz & ",LeaseType as [Lease Type]"
    SQLz = SQLz & ",LeaseTerm as [Lease Term]"
    SQLz = SQLz & ",MilesAllowed As [Miles Allowed]"
    SQLz = SQLz & ",MileCharge as [Mile Charge]"
    SQLz = SQLz & ",BLCD"
    SQLz = SQLz & ",LeaseExp As [Lease Expire]"   'Col AA
    SQLz = SQLz & ",CurrRent As [Curr Rent]"
    SQLz = SQLz & ",CurrRentType as [Rent Type]"
    SQLz = SQLz & ",EOptionSelected AS [EOption Selected]"  'Col AD
    'SQLz = SQLz & ",LeaseTermExtd As [Lease Extd Term]"
    SQLz = SQLz & ",ODORead As [Odometer]"
    SQLz = SQLz & ",ODODate As [Odometer Read Date]"
    SQLz = SQLz & ",ODOSource as [Odometer Source]"  'Col AG
    SQLz = SQLz & ",OffLeaseProj As [Off Lease Proj]"
    SQLz = SQLz & " From vw_ExitReport0"
    SQLz = SQLz & " Where MLOrig='FA' AND ActualOffLease IS NuLL"
    If Nz(Me.SelectClient, 0) > 0 Then
        Dim ClientGrpi As Long
        Dim ClientGroupName As String
        ClientGrpi = Nz(Me.SelectClient.Column(0), 0)
        ClientGroupName = Me.SelectClient.Column(1)
        SQLz = SQLz & " AND ClientGroupID=" & ClientGrpi & " "
        Headerx = Headerx & " - for Client Group: " & ClientGroupName
    End If
    SQLz = SQLz & " Order By Client, Sch,AssetID"
    MessageUser "Open", "Getting Mini Exit Report", "This will take a second or two"
    XLS = "|A,DD|D,BC|E,TX|"
    XLS = XLS & "G,cLabel:End=G:3:Unit" & vbNewLine & "Status|G,BC|H,TC|J,TC|P,TC|"
    XLS = XLS & "Q,cLabel:End=Q:3:Axle" & vbNewLine & "Config|Q,TC|"
    XLS = XLS & "R,cLabel:End=R:3:Unit Config" & vbNewLine & "[UOM]|"
    XLS = XLS & "AC,TC|"
   ' XLS = XLS & "AE,cLabel:End=AE:3:Lease" & vbNewLine & "Ext Term|"
    'XLS = XLS & "AF,cLabel:End=AF:3:Off Lease" & vbNewLine & "Projected|AF,TC|"
    XLS = XLS & "AF,cLabel:End=AF:3:Odometer" & vbNewLine & "Read Date|AF,TC|"
    XLS = XLS & "AG,cLabel:End=AG:3:Odometer" & vbNewLine & "Source|"
    XLS = XLS & "U,$R|V,TC|W,TC|X,Sint|Y,$R|AB,$R|AD,TC|AE,Sint|AG,TC|AH,TC|"
    XLS = XLS & "D,HH|A,W13|I,W21|M,W11|Q,W11|T,W13|AG,W11|AD,W11|AG,W11|AH,W11|"
    Call SendReport2Excel(SQLz, "Units", Headerx, True, XLS, 0, True)
    MessageUser "Close"
End Sub

Private Sub MLInsuranceReport_Click()
Headerx = "Master Lease Insurance Report"
MessageUser "Open", "Getting ML Insurance Report", "This will take a second or two"
    SQLz = "SELECT S.Client"
    SQLz = SQLz & ", MstrLease.MLNo"
    SQLz = SQLz & ", MstrLease.MLDt"
    SQLz = SQLz & ", MstrLease.MLINSAccpt"
    SQLz = SQLz & ", MstrLease.PhysIncAccpt"
    SQLz = SQLz & ", MstrLease.MLInsRec AS InsRecv"
    SQLz = SQLz & ", MstrLease.MLInsExp AS InsExpire"
    SQLz = SQLz & ", MstrLease.PhysInsExp"
    SQLz = SQLz & ", S.AUnits"
    SQLz = SQLz & ", MstrLease.INScomment"
    SQLz = SQLz & ", S.MLID"
    SQLz = SQLz & " FROM (Select ClientShortName As Client, MLID, Count(AssetID) as AUnits "
    SQLz = SQLz & " From vw_SixKeys "
    SQLz = SQLz & " WHERE MLOrig='FA' AND OffLeaseActual IS NULL "
    SQLz = SQLz & " Group BY ClientShortName,MLID)  AS S "
    SQLz = SQLz & " INNER JOIN MstrLease ON S.MLID = MstrLease.MLID"
    SQLz = SQLz & " ORDER BY S.Client;"
    XLS = "|I,BC|I,SintE|K,HH|"
    XLS = XLS & ""
    Call SendReport2Excel(SQLz, "Ins", Headerx, True, XLS, 0, True)
    MessageUser "Close"
End Sub

Private Sub OpeMgtMenu_Click()
 DoCmd.OpenForm "ManagementMenu"
End Sub

Private Sub OutStandingATF_Click()
AddReportLog "OutStandingATF_Click"
DoCmd.OpenReport "OutstandingATFSchedule_KPI_Report", acViewReport
End Sub
Private Sub CompletedATFKPI_Click()
AddReportLog "CompletedATFKPI_Click"
DoCmd.OpenReport "CompletedATFSchedule_KPI_Report", acViewReport
End Sub
Private Sub DealToBLCD_Click()
DoCmd.OpenReport "Deal_to_BLCD_KPI_Report", acViewReport
AddReportLog "DealToBLCD_Click"
End Sub
Private Sub SyndicationLog_Click()
Call MessageUser("Open", "Getting Syndication Master", "This will take a bit")
  SQLz = "Select * From vw_SyndicationBackLog"
  XLS = "|P,BC|S,BC|T,$E|AF,BC|AI,BC|AN,$E|AO,$E|AP,P2C|AQ,$E|AR,P2C|AU,BC|AV,P2C|AW,P2C|AX,P2C|BA,BC|BB,P2C|BC,P2C|BD,P2C|BE,$E|BF,$E|BK,P2C|BM,$E|BN,$E|"
  AddReportLog "SyndicationLog_Click"
  Call SendReport2Excel(SQLz, "Syndication", "Syndication Report " & Date, True, XLS, 0, True)
Call MessageUser("Close")
End Sub

Private Sub YTD0_Click()
    Call SetDateRange("StartDateRange", "EndDateRange", 0)
End Sub
Private Sub YTD1_Click()
    Call SetDateRange("StartDateRange", "EndDateRange", 1)
End Sub
Private Sub YTD2_Click()
    Call SetDateRange("StartDateRange", "EndDateRange", 2)
End Sub
Public Sub SetButtonYears()
    Dim x As Integer
    Dim bName As String
    For x = 0 To 2
        bName = "YTD" & x
        Me(bName).caption = Right(DatePart("yyyy", Date) - x, 4)
    Next x
End Sub
Private Sub SetDateRange(StartField As String, EndField As String, YearBack As Integer)
    Me(StartField) = "7/1/" & DatePart("yyyy", Date) - (YearBack + 1)
    'If YearBack = 0 Then Me(EndField) = Date Else
    Me(EndField) = "6/30/" & DatePart("yyyy", Date) - YearBack
End Sub
