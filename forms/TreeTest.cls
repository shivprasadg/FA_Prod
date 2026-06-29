Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub FindIT_Click()
Call TreeVeiwFindIT(Me.SearchIT)
End Sub

Private Sub Form_Load()
'Set SixTreeView.ImageList = ImageList1
Call SetupTreeview
Call LoadSixTree
End Sub

Private Sub SetupTreeview()
With Me.SixTreeView
.Style = tvwTreelinesPlusMinusPictureText
'tvwTreelinesPlusMinusPictureText
.LineStyle = tvwRootLines
.Indentation = 300
.Appearance = ccFlat
.HideSelection = False
.BorderStyle = ccFixedSingle
.HotTracking = True
.FullRowSelect = True
.CheckBoxes = False
.SingleSel = False
.Sorted = False
.Scroll = True
.LabelEdit = tvwManual
.Font.Name = "Segoe UI"
.Font.Weight = 500
.Font.Size = 8.5

End With
End Sub
Public Function xgetTV() As MSComctlLib.TreeView
  ' Set getTV = Forms(Me.Form.Name).SixTreeView.Object
End Function

Public Sub LoadSixTree(Optional FilterX As String)
   ' On Error GoTo Err_Handler
    Dim SQLz As String
    Dim dbs As DAO.DataBase
    Dim RST As DAO.Recordset
    Dim tv As MSComctlLib.TreeView
    Dim NodX As MSComctlLib.Node
    Dim iNode As Integer
    Dim ClientName, ClientKey As String
    Dim MstName, MstKey, MstPrefix As String
    Dim SchName, SchKey As String
    Dim GrpName, GrpKey As String
    Dim UnitName, UnitKey As String
    Dim ImageX, ImageOpenX As String
    Set tv = getTV(Me.Form.Name)
    tv.Nodes.Clear
    Set dbs = CurrentDb
    SQLz = "SELECT ClientShortName as Client, MLNO, Schedule, ClientID, MLID, SchID, UnitGroup, SGrpID, " & _
             "MLorig, AssetID, UnitDescShortVer & ' (' & Status & ')' AS unitdesc, AssetsInGroup " & _
             "FROM vw_SixKeys "
    Me.FleetCaption.caption = "All Units"
    If FilterX = "FA" Then
      SQLz = SQLz & "WHERE MLOrig ='FA' "
      Me.FleetCaption.caption = "FA Lease Only"
    End If
    If FilterX = "Track" Then
    SQLz = SQLz & "WHERE MLOrig <> 'FA' "
    Me.FleetCaption.caption = "Tracking Only"
    End If
    SQLz = SQLz & "ORDER BY ClientShortName, MLorig, MLNO, SchID Desc, SGrpID;"
    Set RST = dbs.OpenRecordset(SQLz, dbOpenSnapshot)
    Set NodX = tv.Nodes.Add(, , ClientKey, ClientName)

    'Me.SixTreeView.Nodes.Add Text:="Clients", Key:="ClientID"
    If Not RST.BOF Then RST.MoveFirst
    Do Until RST.EOF
         ' Add Client nodes
        If ClientName <> RST!Client Then
            ClientKey = RST!ClientID & "|"
            ClientName = ProperCase(RST!Client, 0)
            Set NodX = tv.Nodes.Add(, , ClientKey, ClientName, "Client", "ClientOpen")
            NodX.Tag = "Client:" & RST!ClientID
            MstName = ""
        End If


        ' Add Store node
        If MstName <> RST!MLNo Then
            MstKey = ClientKey & RST!MLID & "|"
            MstName = RST!MLNo
            If RST!MLOrig = "FA" Then
            MstPrefix = "Lease "
            ImageX = "Lease"
            ImageOpenX = "LeaseOpen"
            Else
            MstPrefix = "Trk "
            ImageX = "Tracking"
            ImageOpenX = "TrackingOpen"
            End If
            Set NodX = tv.Nodes.Add(ClientKey, tvwChild, MstKey, MstName, ImageX, ImageOpenX)
            NodX.Tag = ImageX & ":" & RST!MLID
            SchName = ""
        End If

        ' Add Author node
        On Error Resume Next
        If SchName <> RST!Schedule Then
            SchKey = MstKey & RST!SchID & "|"
            SchName = RST!Schedule
            ImageX = "Schedule"
            Set NodX = tv.Nodes.Add(MstKey, tvwChild, SchKey, "Sch " & SchName, ImageX, "FolderOpen")
            NodX.Tag = "Schedule:" & RST!SchID
            GrpName = ""
        End If
        On Error GoTo 0

        On Error Resume Next
        If GrpName <> RST!UnitGroup & " | " & RST!AssetsInGroup & " units" Then
            GrpKey = SchKey & RST!SGrpID & "|"
            GrpName = RST!UnitGroup & " | " & RST!AssetsInGroup & " units"
            ImageX = "GroupName"
            Set NodX = tv.Nodes.Add(SchKey, tvwChild, GrpKey, "Grp " & GrpName, ImageX, "GroupNameOpen")
            NodX.Tag = "Group:" & RST!SGrpID
            UnitName = ""
           End If
        On Error GoTo 0

        On Error Resume Next
        If UnitName <> RST!AssetID & " " & RST!UnitDesc Then
            UnitKey = GrpKey & RST!AssetID & "|"
            UnitName = RST!AssetID & " " & RST!UnitDesc
            ImageX = "Truck"
            If RST!UnitDesc Like "*Day*" Then ImageX = "DayCab"
            If RST!UnitDesc Like "*Sleep*" Then ImageX = "Sleeper"
            If RST!UnitDesc Like "*Trai*" Then ImageX = "Trailer"
            Set NodX = tv.Nodes.Add(GrpKey, tvwChild, UnitKey, UnitName, ImageX, ImageX)
            NodX.Tag = "Unit:" & RST!AssetID
        End If
        On Error GoTo 0
MoveOn:

        RST.MoveNext
    Loop

    ' Initialize ListView to empty recordset.
    'LoadMyListView ("")

Exit_Here:
    Set dbs = Nothing
    Set RST = Nothing
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbCritical, "ERROR"
    'Resume Next
End Sub

Private Sub LoadMyListView(ByVal au_id As String)
    On Error GoTo Err_Handler
    Dim SQLz As String
    Dim dbs As DAO.DataBase
    Dim RST As DAO.Recordset
    Dim lstItem As ListItem
    Dim strItem As String
    Dim strKey As String
    Dim intCount As Integer


    ' Remove all existing items
    While Me!lvwMyListView.ListItems.Count > 0
        Me!lvwMyListView.ListItems.Remove (1)
    Wend

    'open recordset of filtered Title/Authors
    Set dbs = CurrentDb
    SQLz = "SELECT title, title_id FROM qryTitleAuthor WHERE [au_id]='" & au_id & "'"
    Set RST = dbs.OpenRecordset(SQLz, dbOpenSnapshot)

    'add each entry in the rs to the listview
    If Not RST.BOF Then RST.MoveFirst
    Do Until RST.EOF
        strItem = RST!Title
        strKey = "ID=" & RST!title_id

        Set lstItem = Me!lvwMyListView.ListItems.Add(, strKey, strItem)

        intCount = intCount + 1
        RST.MoveNext
    Loop

    If intCount > 0 Then Me!lvwMyListView.ListItems(1).Selected = True

Exit_Here:
    Set dbs = Nothing
    Set RST = Nothing
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbCritical, "ERROR"
    Resume Next
End Sub

Private Sub ReloadTree_Click()
Call SetupTreeview
Call LoadSixTree
End Sub
Private Sub SixTreeView_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Long, ByVal Y As Long)
    Dim NodX   As Node
    Set NodX = getTV(Me.Form.Name).HitTest(x, Y)
    If Button = eMouse.Rightclick Then
        If NodX Is Nothing Then
            Call RightClickEmptySpace(Me.Form.Name)
        Else
            Call RightClickNode(Me.Form.Name, NodX.Tag)
            'MsgBox nodX.Text & vbNewLine & nodX.Tag
        End If
    End If

    '  If Button = acRightButton Then
    '    CommandBars("MyTreeviewContextMenu").ShowPopup
    '  End If
End Sub
