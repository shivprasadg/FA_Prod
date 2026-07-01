Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim ClientIDx  As Long
Public PrevValue As Variant
Public pField  As String
Dim RST        As DAO.Recordset
Dim Committ As String

Private Sub Axle_DblClick(Cancel As Integer)
    If Me.Axle & "test" <> "test" Then Call SetItems(Me.Axle, "Axle", "Axle", "Unit")
End Sub

Public Sub FAIDList_Click()
    Me.Form.RecordSource = "SELECT * FROM ATF_Schedule_Group Where ATFID=" & Me.FAIDList & ";"
    SQLz = "SELECT " & _
        "ATFGroupSpecID, " & _
        "ATFGroupSpecDesc As [Group]" & _
        ",Nz([UnitsInGroup],0) AS Units" & _
        ",ATFID " & _
        "From ATF_Schedule_Group " & _
        "Group By ATFGroupSpecID," & _
        "ATFGroupSpecDesc, Nz([UnitsInGroup],0), ATFID " & _
        "HAVING Nz([UnitsInGroup],0) >0 AND ATFID=" & Me.FAIDList & " " & _
        "Order By ATFGroupSpecDesc ;"
    Me.GroupList.RowSource = SQLz
End Sub

Private Sub SetComboBoxes()
    SQLz = "SELECT ATFGroupSpecID, ATFGroupSpecDesc, ATFID, ATFSchSpecID " & _
        "From ATF_Schedule_Group " & _
        "WHERE ATFSchSpecID = " & Me.ATFSchSpecID & " " & _
        "GROUP BY ATFGroupSpecID, ATFGroupSpecDesc, ATFID, ATFSchSpecID;"
    Me.GroupsList.RowSource = SQLz


    SQLz = "SELECT ATFChildSpecID, ATFChildSpecDesc, ATFID, ATFUnitSpecID " & _
        "FROM ATF_Schedule_Group " & _
        "GROUP BY ATFChildSpecID, ATFChildSpecDesc, ATFID, ATFUnitSpecID " & _
        "HAVING ATFUnitSpecID = " & Me.ATFUnitSpecID & " ORDER BY ATFChildSpecID;"
    Me.ChildList.RowSource = SQLz

    SQLz = "SELECT LocationID, Locationsn AS Location, LocationType, LAddrLine1 AS Adress, LState AS St, LZip AS Zip, LCity AS City " & _
        "From Locations WHERE LocationType<>'billing' AND LocObsolete=No AND clientRef=" & Nz(Me.ClientID, 0) & " " & _
        "ORDER BY LState, LCity, LocationName;"
    Me.DeliveryLocationID.RowSource = SQLz
    Me.RegTitleLocationID.RowSource = SQLz


    ClientIDx = Nz(DLookup("LesseeID", "ATF", "ATFID=" & Nz(Me.ATFID, 0)), 0)
    SQLz = "SELECT LocationID, Nz([LocationName],[Locationsn]) & ' - ' & Nz([LAddrLine1],[LAddrLine2]) & ' ' & [LCity] & ', ' & [LState] & ' ' & [LZip] AS [Billing Location], " & _
        "LState As State, LZip, LocationType, LocationName, LCity FROM Locations " & _
        "WHERE (LocationType ='Billing' Or LocationType='Garage') AND Locations.LocObsolete=No AND ClientRef=" & Nz(Me.ClientID, 0) & " ORDER BY Locations.LState;"
    Me.ScheduleBillingID.RowSource = SQLz
End Sub

Private Sub Form_Open(Cancel As Integer)
    Set Engine = New clsFieldEngine
End Sub
Private Sub Form_Load()
    Me.SearchClient = ""
    Me.SearchFAID = ""
    Me.SearchUnitSpecID = ""
    SQLz = "SELECT A.ATFID, C.ClientShNm AS Client, A.SchFAID AS FAID, C.clientID"
    SQLz = SQLz & " FROM ATF_Schedule_Group AS A INNER JOIN Clients AS C ON A.ClientID = C.ClientID"
    SQLz = SQLz & " GROUP BY A.ATFID, C.ClientShNm, A.SchFAID, C.ClientID"
    SQLz = SQLz & " ORDER BY C.ClientShNm, A.SchFAID DESC;"
    Me.FAIDList.RowSource = SQLz
End Sub
Private Sub Form_Current()
    Call Refresh_My_Buttons
    If FaUserName = "ljohnson" Then
        Dim Ccrl As Control
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Spec" Then
                Me(Ccrl.Name).Locked = False
            End If
        Next Ccrl
    End If
End Sub

Public Sub GroupList_Click()
    SQLz = "Select * From ATF_Schedule_Group Where ATFGroupSpecID = " & Me.GroupList
    Me.Form.RecordSource = SQLz
    Me.ChildList.Requery
End Sub

Private Sub LeaseExtOption1Term_AfterUpdate()

End Sub

Private Sub LeaseExtOption1Term_GotFocus()
    PrevValue = Me.LeaseExtOption1Term
    pField = "LeaseExtOption1Term"
End Sub

Private Sub ScheduleBillingID_GotFocus()
    Me.ScheduleBillingID.Requery
End Sub

Private Sub Refresh_My_Buttons()
    Dim ctrl   As Control
    On Error GoTo Refresh_My_Buttons_Error

    For Each ctrl In Me.Controls
        If ctrl.ControlType = acCommandButton Then ctrl.Requery
    Next ctrl
    On Error GoTo 0
    Exit Sub

Refresh_My_Buttons_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Refresh_My_Buttons of VBA Document Form_ATF_ATFSpec"
End Sub

Private Sub MakeChildPartSpec_Click()
    If Nz(Me.NumOfChildSpecs, 0) = 0 Then
        MsgBox "You must first enter the number of Child Part Specs that are within this Unit Spec.", , AtlaasVersion
        Exit Sub
    End If
    Dim xi, xCount As Integer
    Dim LastExitstingSpec As Integer
    Dim DescX, TempStr As String

    xCount = DCount("[ATFChildSpecID]", "ATFChildPartSpec", "ATFUnitSpecID=" & Me.ATFUnitSpecID)
    If xCount >= Me.NumOfChildSpecs Then
        MsgBox "All of the Child Part Specs within the Unit Spec have been created."
        Exit Sub
    Else
        If xCount > 0 Then LastExitstingSpec = CInt(Right(DMax("ATFChildSpecDesc", "ATFChildPartSpec", "ATFUnitSpecID=" & Me.ATFUnitSpecID), 1))
        If (Me.NumOfChildSpecs - LastExitstingSpec) > 1 Then TempStr = " New Child Part Specs" Else TempStr = " New Unit Spec"
        If LastExitstingSpec = 1 Then TempStr = TempStr & vbNewLine & "(" & LastExitstingSpec & " Child Part Spec alreadly exists in Atlass Core)" & vbNewLine
        If LastExitstingSpec > 1 Then TempStr = TempStr & vbNewLine & "(" & LastExitstingSpec & " Child Part Specs alreadly exists in Atlass Core)" & vbNewLine
        If AskToDoTask("You're about to Insert " & (Me.NumOfChildSpecs - LastExitstingSpec) & TempStr) = True Then
            For xi = (1 + LastExitstingSpec) To Me.NumOfChildSpecs
                SQLz = ""
                DescX = "Child Spec " & xi
                SQLz = SQLz & "INSERT INTO ATFChildPartSpec "
                SQLz = SQLz & "([ATFID]"
                SQLz = SQLz & ",[ATFUnitSpecID]"
                SQLz = SQLz & ",[ATFChildSpecDesc]"
                SQLz = SQLz & ",[ChildVendorID]"
                SQLz = SQLz & ",[ChildPartYear]"
                SQLz = SQLz & ",[ChildMake]"
                SQLz = SQLz & ",[ChildClass]"
                SQLz = SQLz & ",[ChildType]"
                SQLz = SQLz & ",[ChildPartCost]"
                SQLz = SQLz & ") "

                SQLz = SQLz & "Values ("
                SQLz = SQLz & Me.ATFID
                SQLz = SQLz & "," & Me.ATFUnitSpecID
                SQLz = SQLz & ",'" & DescX & "'"
                SQLz = SQLz & ",39 "
                SQLz = SQLz & "," & DatePart("YYYY", Date) + 1
                SQLz = SQLz & ",'TBD'"
                SQLz = SQLz & ",'TBD'"
                SQLz = SQLz & ",'TBD'"
                SQLz = SQLz & ",0"
                SQLz = SQLz & ");"
                DoCmd.RunSQL (SQLz)
            Next xi
        End If
    End If

End Sub
Private Sub MakeChildPart_Click()
    Dim xi     As Integer
    SQLz = ""
    SQLz = SQLz & "INSERT INTO childparts ("
    SQLz = SQLz & "ATFUnitSpecID"
    SQLz = SQLz & ", ATFChildSpecID"
    SQLz = SQLz & ", UnitREF"
    SQLz = SQLz & ", ChildVendorREF"
    SQLz = SQLz & ", ChildMake"
    SQLz = SQLz & ", ChildClass"
    SQLz = SQLz & ", ChildType"
    SQLz = SQLz & ", childFinalCost "
    SQLz = SQLz & ") SELECT "
    SQLz = SQLz & " ATFUnitSpecID "
    SQLz = SQLz & ",ATFChildSpecID "
    SQLz = SQLz & "," & DLookup("UnitID", "Units", "")
    SQLz = SQLz & ",ChildVendorID "
    SQLz = SQLz & ",ChildMake "
    SQLz = SQLz & ",ChildClass "
    SQLz = SQLz & ",ChildType "
    SQLz = SQLz & ",ChildPartCost "
    SQLz = SQLz & " FROM ATFChildPartSpec WHERE ATFUnitSpecID=" & Me.ATFUnitSpecID
    DoCmd.RunSQL (SQLz)
End Sub

Private Sub SearchClient_Change()
    Me.SearchFAID = ""
    SQLz = "SELECT A.ATFID, C.ClientShNm AS Client, A.SchFAID AS FAID, C.clientID"
    SQLz = SQLz & " FROM ATF_Schedule_Group AS A INNER JOIN Clients AS C ON A.ClientID = C.ClientID"
    SQLz = SQLz & " WHERE (((C.ClientShNm) Like '*" & Me.SearchClient.Text & "*'))"
    SQLz = SQLz & " GROUP BY A.ATFID, C.ClientShNm, A.SchFAID, C.ClientID"
    SQLz = SQLz & " ORDER BY C.ClientShNm, A.SchFAID DESC;"

    Me.FAIDList.RowSource = SQLz
End Sub

Public Sub SearchFAID_Change()
    Me.SearchClient = ""
    SQLz = "SELECT A.ATFID, C.ClientShNm AS Client, A.SchFAID AS FAID, C.clientID"
    SQLz = SQLz & " FROM ATF_Schedule_Group AS A INNER JOIN Clients AS C ON A.ClientID = C.ClientID"
    SQLz = SQLz & " WHERE (((A.SchFAID) Like '*" & Me.SearchFAID.Text & "*'))"
    SQLz = SQLz & " GROUP BY A.ATFID, C.ClientShNm, A.SchFAID, C.ClientID"
    SQLz = SQLz & " ORDER BY C.ClientShNm, A.SchFAID DESC;"
    Me.FAIDList.RowSource = SQLz
End Sub

Public Sub GetUnitSpecID_Click()
    If Me.SearchUnitSpecID & "test" = "test" Then
        MsgBox "There must be a UnitSpec ID Entered before you can Ssearch it", , "No Results"
        Exit Sub
    End If

    SQLz = "Select S.SchFAID, S.ATFID, S.ATFSchSpecID, G.ATFGroupSpecID, U.ATFUnitSpecID"
    SQLz = SQLz & " FROM ((ATFScheduleSpec As S"
    SQLz = SQLz & " Inner Join ATFGroupSpec AS G ON S.ATFSchSpecID = G.ATFSchSpecID)"
    SQLz = SQLz & " Inner Join ATFUnitSpec AS U ON G.ATFGroupSpecID = U.ATFGroupSpecID)"
    SQLz = SQLz & " WHERE ATFUnitSpecID =" & Me.SearchUnitSpecID

    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        Me.SearchFAID = RST!SchFAID
        ' Call SearchFAID_Change
        Me.FAIDList = RST!ATFID
        Call FAIDList_Click
        Me.GroupList = RST!ATFGroupSpecID
        Call GroupList_Click
    Else
        MsgBox "There isn't a UnitSpecID Match for: " & Me.SearchUnitSpecID, vbCritical, "Get some glasses!"
    End If
End Sub
Public Function UndoChange(changeLabel As String)
    Me.Undo
    MsgBox "There wasn't a " & changeLabel & " Change", vbInformation + vbOK, "Like It Never Happened"
    PrevValue = Null
    pField = ""
    Me.ChildMake.SetFocus
End Function

Private Sub PowerType_DblClick(Cancel As Integer)
    If SetLPAFItems(Me.PowerType, "PowerType", "Power Type", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Power Type")
End Sub

Private Sub TransmissionType_DblClick(Cancel As Integer)
    If SetLPAFItems(Me.TransmissionType, "UnitTransmissionType", "Transmission", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Transmission")
End Sub

Private Sub UnitAxleProperty_DblClick(Cancel As Integer)
    If SetLPAFItems(Me.UnitAxleProperty, "UnitAxleProperty", "Axle Prop", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("UnitAxleProperty")
End Sub

Private Sub UnitConfiguration_DblClick(Cancel As Integer)
    If SetLPAFItems(Me.UnitConfiguration, "UnitConfiguration", "Unit Config", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Unit Config")
End Sub

Private Sub UnitFinalCost_AfterUpdate()
    If PrevValue <> Me.UnitFinalCost And pField = "UnitFinalCost" Then
        If ChangeLPAFItems(Me.UnitFinalCost, PrevValue, pField, "Unit Cost", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Unit Cost")
    End If
End Sub

Private Sub UnitFinalCost_GotFocus()
    PrevValue = Me.UnitFinalCost
    pField = "UnitFinalCost"
End Sub

Private Sub UnitMake_Click()

End Sub

Private Sub UnitUOM_DblClick(Cancel As Integer)
    If SetLPAFItems(Me.UnitUOM, "UnitUOM", "Unit UOM", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Unit Uom")
End Sub
Private Sub UnitDescriptor_DblClick(Cancel As Integer)
    If SetLPAFItems(Me.UnitDescriptor, "UnitDescriptor", "Unit Descriptor", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Unit Descriptor")
End Sub
Private Sub EngineType_DblClick(Cancel As Integer)
    If SetLPAFItems(Me.EngineType, "UnitEngineType", "Engine Type", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Engine Type")
End Sub
Private Sub UnitSubType_AfterUpdate()
    If PrevValue <> Me.UnitSubType And pField = "UnitSubType" Then
        If ChangeLPAFItems(Me.UnitSubType, PrevValue, pField, "Unit SubType", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Unit SubType")
    End If
End Sub
Private Sub UnitSubType_GotFocus()
    PrevValue = Me.UnitSubType
    pField = "UnitSubType"
End Sub

Private Sub UnitType_AfterUpdate()
    If PrevValue <> Me.UnitType And pField = "UnitType" Then
        If ChangeLPAFItems(Me.UnitType, PrevValue, pField, "UnitType", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("UnitType")
    End If
End Sub
Private Sub UnitType_GotFocus()
    PrevValue = Me.UnitType
    pField = "UnitType"
End Sub

Private Sub UNITSRENT_GotFocus()
    PrevValue = Me.UNITSRENT
    pField = "UnitSRent"
End Sub
Private Sub UNITSRENT_AfterUpdate()
    If PrevValue <> Me.UNITSRENT And pField = "UnitSRent" Then
        If ChangeLPAFItems(Me.UNITSRENT, PrevValue, pField, "Sch Rent", "Unit", Me.ATFUnitSpecID) = 2 Then UndoChange ("Sch Rent")
    End If
End Sub

Private Sub UnitType_Click()
    Me.UnitSubType.Requery
    Call SetUOM
End Sub
Private Sub UnitSubType_Click()
    Me.UnitConfiguration.Requery
    Call SetUOM
End Sub

Private Sub UnitUOM_GotFocus()
    PrevValue = Me.UnitUOM
    pField = "UnitUOM"
    Call SetUOM
End Sub

Private Sub SetUOM()
    ' Set UOM Control Tip and ComboBox
    Dim UomX   As Integer
    Me.UnitUOM.LimitToList = True
    Me.UnitUOM.ColumnWidths = "0;0.8"
    If Me.UnitSubType = "Sleeper" Then UomX = 12
    If Me.UnitType = "Straight Truck" Then UomX = 20
    If Me.UnitType = "Trailer" Then UomX = 30
    If Me.UnitSubType = "Auto Hauler" Then UomX = 36
    If Me.UnitSubType = "Tank" Then UomX = 37
    If Me.UnitSubType = "Step Van" Then UomX = 20
    If Me.UnitType = "Straight Truck" And Me.UnitSubType = "Tank" Then UomX = 27
    If UomX <> 0 Then
        Dim ix As Integer
        Dim ListX As String
        Me.UnitUOM.ControlTipText = "A " & Me.UnitSubType & " " & Me.UnitType & "'s UOM Range is: " & DLookup("UOMLabel", "vw_UnitType_Matrix", "UOMGroupID = " & UomX)
        'Sleeper UOM
        If UomX = 12 Then
            For ix = 37 To 80
                ListX = ListX & ix & ";" & ix & "in;"
            Next ix
        End If
        'Straight Truck UOM
        If UomX = 20 Then
            For ix = 12 To 38
                ListX = ListX & ix & ";" & ix & "ft;"
            Next ix
        End If
        'Trailer UOM
        If UomX = 30 Then
            For ix = 28 To 55
                ListX = ListX & ix & ";" & ix & "ft;"
            Next ix
        End If
        'Tank UOM
        If UomX = 27 Or UomX = 47 Then
            Me.UnitUOM.ColumnWidths = "0.8;0"
            Me.UnitUOM.LimitToList = False
            For ix = 1000 To 11000 Step 500
                ListX = ListX & ix & ";" & ix & ";"
            Next ix
        End If
        'Auto Hauler UOM
        If UomX = 36 Then
            For ix = 5 To 11
                ListX = ListX & ix & ";" & ProperCase(SayNumber(ix), 0) & " Car;"
            Next ix
        End If

        Me.UnitUOM.RowSource = ListX
        Me.UnitUOM.Enabled = True
    Else
        Me.UnitUOM.RowSource = ""
        Me.UnitUOM.ControlTipText = "A " & Me.UnitSubType & " " & Me.UnitType & " doesn't have a UOM"
    End If
End Sub
