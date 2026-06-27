Attribute VB_Name = "TreeView_Functions"
Option Compare Database
Option Explicit
Public TreeForm As String
Public NodeTypex As String
Public CallingForm As String
Public LKeyID As Long
Public tv     As TreeView
Public NodX   As Node
Public NodChildX As Node

Function MakeATreeNode(SQLx, NodeNodTagX As String)

End Function

Public Sub PopUpNodeDetails()
  With Forms!TreeTest.SixTreeView.SelectedItem
  Dim Ni As Integer
  Dim KeyX As String
  KeyX = .Key
  Ni = CountStringOccurances(KeyX, "|")
  If Ni = 1 Then
  KeyX = (Split(.Key, "|")(0))
  DoCmd.close acForm, "ClientsView"
  DoCmd.OpenForm "ClientsView", , , , , , CInt(KeyX)
  End If
  If Ni = 2 Then
  KeyX = (Split(.Key, "|")(1))
  KeyX = (Split(KeyX, "|")(0))
  DoCmd.close acForm, "MstrLease"
  DoCmd.OpenForm "MstrLease", acNormal, , "MLID=" & CInt(KeyX)
  End If
  If Ni = 3 Then
  KeyX = (Split(.Key, "|")(2))
  KeyX = (Split(KeyX, "|")(0))
  DoCmd.close acForm, "Schedule_Form"
  DoCmd.OpenForm "Schedule_Form", acNormal, , "SchID=" & CInt(KeyX)
  End If
  If Ni = 4 Then
  KeyX = (Split(.Key, "|")(3))
  KeyX = (Split(KeyX, "|")(0))
  DoCmd.close acForm, "GroupForm"
  DoCmd.OpenForm "GroupForm", acNormal, , "SGrpID=" & CInt(KeyX)
  End If
  If Ni = 5 Then
  KeyX = (Split(.Key, "|")(4))
  KeyX = (Split(KeyX, "|")(0))
  'DoCmd.OpenForm "UnitsForm", acNormal, , "UnitID=" & CLng(KeyX)
  End If
   ' MsgBox "Text=" & .Text & vbCrLf & _
   '   "Key=" & .Key & vbCrLf & _
   '   "Tag=" & .Tag
  End With
End Sub

Public Sub PopUpNodeDelete()
  If MsgBox("Delete node " & Forms!frmTreeviewDemo.tvDetails.SelectedItem.Text & "?", _
    Buttons:=vbOKCancel) = vbOK Then

    MsgBox "Node to be deleted"

    ' delete code goes here
    ' usually as well as deleting the node there's some data to be deleted from tables as well
  Else
    MsgBox "Delete cancelled"
  End If

End Sub

Public Sub SetUpContextMenu()
  ' requires a reference to the Microsoft office object library

  On Error Resume Next ' ignore error if command bar does not exist to be deleted
  CommandBars("MyTreeviewContextMenu").Delete
  On Error GoTo 0

  With CommandBars.Add(Name:="MyTreeviewContextMenu", Position:=msoBarPopup)
    With .Controls.Add(Type:=msoControlButton)
      .caption = "Open Form View"
      .OnAction = "PopUpNodeDetails"
    End With
    With .Controls.Add(Type:=msoControlButton)
      .caption = "Delete Node"
      .OnAction = "PopUpNodeDelete"
    End With
  End With
End Sub

Function CountStringOccurances(strStringToCheck As String, strValueToCheck As String) As Integer
'Purpose: Counts the number of times a string appears in another string.
On Error GoTo ErrorMessage
    Dim intStringPosition As Integer
    Dim intCursorPosition As Integer
    Dim i As Integer
    CountStringOccurances = 0
    intCursorPosition = 1
    For i = 0 To Len(strStringToCheck)
        intStringPosition = InStr(intCursorPosition, strStringToCheck, strValueToCheck)
        If intStringPosition = 0 Then
            Exit Function
        Else
            CountStringOccurances = CountStringOccurances + 1
            intCursorPosition = intStringPosition + Len(strValueToCheck)
        End If
    Next i
    Exit Function
ErrorMessage:
MsgBox "The database has generated an error. Please contact the database administrator, quoting the following error message: '" & Err.Description & "'", vbCritical, "Database Error"
End
End Function
Public Function getTV(formName As String) As MSComctlLib.TreeView
   Set getTV = Forms(formName).SixTreeView.Object
    TreeForm = formName
    CallingForm = formName
End Function
Public Function getImageList() As MSComctlLib.ImageList
   Set getImageList = Forms(CallingForm).ImageList3.Object
End Function
Public Sub RightClickEmptySpace(FormNameX As String)
   Dim cmdBAR As CommandBar
   TreeForm = FormNameX
   Set cmdBAR = CommandBars.Add(, msoBarPopup, False, True)

   Dim cmdExpandAll As CommandBarButton
   Set cmdExpandAll = cmdBAR.Controls.Add(msoControlButton)
   cmdExpandAll.caption = "Expand All"
   cmdExpandAll.OnAction = "ExpandAll"
   cmdExpandAll.Picture = getImageList.ListImages("ExpandAll").Picture

   Dim cmdContractAll As CommandBarButton
   Set cmdContractAll = cmdBAR.Controls.Add(msoControlButton)
   cmdContractAll.caption = "Contract All"
   cmdContractAll.OnAction = "ContractAll"
   cmdContractAll.Picture = getImageList.ListImages("ContractAll").Picture

   Dim cmdAllFleets As CommandBarButton
   Set cmdAllFleets = cmdBAR.Controls.Add(msoControlButton)
   cmdAllFleets.caption = "Show All Units"
   cmdAllFleets.OnAction = "ShowAllFleets"
   cmdAllFleets.Picture = getImageList.ListImages("FilterClear").Picture

   Dim cmdJustFA As CommandBarButton
   Set cmdJustFA = cmdBAR.Controls.Add(msoControlButton)
   cmdJustFA.caption = "Show FA Only"
   cmdJustFA.OnAction = "ShowFAOnly"
   cmdJustFA.Picture = getImageList.ListImages("Filter").Picture

   Dim cmdJustBUnits As CommandBarButton
   Set cmdJustBUnits = cmdBAR.Controls.Add(msoControlButton)
   cmdJustBUnits.caption = "B Units Only"
   cmdJustBUnits.OnAction = "ShowBUnitsOnly"
   cmdJustBUnits.Picture = getImageList.ListImages("Filter").Picture

   Dim cmdJustTrack As CommandBarButton
   Set cmdJustTrack = cmdBAR.Controls.Add(msoControlButton)
   cmdJustTrack.caption = "Show Tracking Only"
   cmdJustTrack.OnAction = "ShowTrackOnly"
   cmdJustTrack.Picture = getImageList.ListImages("Filter").Picture

   Dim cmdJustTR As CommandBarButton
   Set cmdJustTR = cmdBAR.Controls.Add(msoControlButton)
   cmdJustTR.caption = "TR Units Only"
   cmdJustTR.OnAction = "ShowTROnly"
   cmdJustTR.Picture = getImageList.ListImages("Filter").Picture

   cmdBAR.ShowPopup

'Cleanup
   Set cmdBAR = Nothing
   Set cmdExpandAll = Nothing
   Set cmdContractAll = Nothing

End Sub
Public Sub RightClickNode(FormNameN As String, Optional NodeInfo As String)
   Dim cmdBAR As CommandBar
   TreeForm = FormNameN
   Set cmdBAR = CommandBars.Add(, msoBarPopup, False, True)
   If NodeInfo <> "" Then
   NodeTypex = Split(NodeInfo, ":")(0)
  ' If IsNumeric(NodeTypex) Then
   LKeyID = CLng(Nz(Split(NodeInfo, ":")(1), 0))
   Else
   NodeTypex = ""
   End If
   Dim cmdOpenFrom As CommandBarButton
   Set cmdOpenFrom = cmdBAR.Controls.Add(msoControlButton)
   cmdOpenFrom.caption = "Open " & NodeTypex & " Form"
   cmdOpenFrom.OnAction = "OpenForm"
   cmdOpenFrom.Picture = getImageList.ListImages("GetInfo").Picture

   If NodeTypex = "Group" Then
   Dim cmdMassMod As CommandBarButton
   Set cmdMassMod = cmdBAR.Controls.Add(msoControlButton)
   cmdMassMod.caption = "Open Group In Mass_Mod"
   cmdMassMod.OnAction = "OpenMassModForm"
   cmdMassMod.Picture = getImageList.ListImages("MassMod").Picture
   End If

   Dim cmdExpandMe As CommandBarButton
   Set cmdExpandMe = cmdBAR.Controls.Add(msoControlButton)
   cmdExpandMe.caption = "Expand To Group"
   cmdExpandMe.OnAction = "ExpandMe"
   cmdExpandMe.Picture = getImageList.ListImages("ExpandAll").Picture

   Dim cmdExpandMeAll As CommandBarButton
   Set cmdExpandMeAll = cmdBAR.Controls.Add(msoControlButton)
   cmdExpandMeAll.caption = "Expand To Unit"
   cmdExpandMeAll.OnAction = "ExpandMeAll"
   cmdExpandMeAll.Picture = getImageList.ListImages("ExpandAll").Picture

   Dim cmdContractAll As CommandBarButton
   Set cmdContractAll = cmdBAR.Controls.Add(msoControlButton)
   cmdContractAll.caption = "Contract All"
   cmdContractAll.OnAction = "ContractAll"
   cmdContractAll.Picture = getImageList.ListImages("ContractAll").Picture
   cmdBAR.ShowPopup

'Cleanup
   Set cmdBAR = Nothing
   Set cmdOpenFrom = Nothing
   Set cmdContractAll = Nothing
End Sub
Public Sub OpenForm()
   Dim tv As TreeView
   Dim NodX As Node
   Set tv = getTV(TreeForm)
   If NodeTypex = "Client" Then DoCmd.OpenForm "ClientsView", , , , , , LKeyID
   If NodeTypex = "Lease" Then DoCmd.OpenForm "MstrLease", acNormal, , "MLID=" & LKeyID
   If NodeTypex = "Tracking" Then DoCmd.OpenForm "MstrLease", acNormal, , "MLID=" & LKeyID
   If NodeTypex = "Schedule" Then DoCmd.OpenForm "Schedule_Form", acNormal, , "SchID=" & LKeyID
   If NodeTypex = "Group" Then DoCmd.OpenForm "GroupForm", acNormal, , "SGrpID=" & LKeyID
  ' If NodeTypex = "Unit" Then 'DoCmd.OpenForm "UnitsForm", acNormal, , "UnitID=" & LKeyID

'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub
Public Sub OpenMassModForm()
   Dim tv As TreeView
   Dim NodX As Node
   Set tv = getTV(TreeForm)
    DoCmd.close acForm, "BulkChange_form"
    DoCmd.OpenForm "BulkChange_form"
    Forms!BulkChange_Form.ClientList = DLookup("ClientId", "vw_SixKeys", "sGrpID=" & LKeyID)
    Form_BulkChange_Form.ClientList_Click
    Forms!BulkChange_Form.LeaseList = DLookup("MLID", "vw_SixKeys", "sGrpID=" & LKeyID)
    Form_BulkChange_Form.Leaselist_Click
    Forms!BulkChange_Form.ScheduleList = DLookup("SchID", "vw_SixKeys", "sGrpID=" & LKeyID)
    Form_BulkChange_Form.ScheduleList_Click
    Forms!BulkChange_Form.GroupList = LKeyID
    Form_BulkChange_Form.GroupList_Click
    Form_BulkChange_Form.GetAll_Click
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub
Public Sub ShowFAOnly()
    Set tv = getTV(TreeForm)
    Call Forms(TreeForm).LoadSixTree("FA")
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub
Public Sub ShowBUnitsOnly()
    Set tv = getTV(TreeForm)
    Call Forms(TreeForm).LoadSixTree("BUnits")
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub
Public Sub ShowTrackOnly()
    Set tv = getTV(TreeForm)
    Call Forms(TreeForm).LoadSixTree("Track")
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub
Public Sub ShowAllFleets()
    Set tv = getTV(TreeForm)
    Call Forms(TreeForm).LoadSixTree
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub
Public Sub ShowTROnly()
    Set tv = getTV(TreeForm)
    Call Forms(TreeForm).LoadSixTree("TRUnits")
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub
Public Sub ExpandAll()
   Set tv = getTV(TreeForm)
   For Each NodX In tv.Nodes
      If NodX.Expanded = False Then
         NodX.Expanded = True
      End If
   Next
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub
Public Sub ExpandMe()
    Dim tv     As TreeView
    Dim NodX   As Node
    Dim NodChildX As Node
    Set tv = getTV(TreeForm)
    Dim i, x     As Long
    tv.SelectedItem.Expanded = True
  Call ExpandChildren(tv.SelectedItem, False)
tv.SelectedItem.EnsureVisible
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
    Exit Sub
End Sub
Public Sub ExpandMeAll()
    Set tv = getTV(TreeForm)
    Dim i, x     As Long
    tv.SelectedItem.Expanded = True
  Call ExpandChildren(tv.SelectedItem, True)
tv.SelectedItem.EnsureVisible
    'cleanup
    Set NodX = Nothing
    Set tv = Nothing
    Set NodChildX = Nothing
    Exit Sub
End Sub
Private Sub ExpandChildren(PNode As Node, toUnit As Boolean)
    Dim i      As Integer
    Dim ThisNode As Node
    If PNode.Children <> 0 Then
        Set ThisNode = PNode.Child
        For i = 1 To PNode.Children
            If toUnit = False And ThisNode Like "*Unit*" Then ThisNode.Expanded = False Else ThisNode.Expanded = True
            Call ExpandChildren(ThisNode, toUnit)
            Set ThisNode = ThisNode.Next
        Next
    End If
End Sub

Public Sub ContractAll()
   Set tv = getTV(TreeForm)
   For Each NodX In tv.Nodes
      If NodX.Expanded = True Then
         NodX.Expanded = False
      End If
   Next
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
End Sub

Private Sub MakeWorkTable()
Dim SQLz As String
    DoCmd.SetWarnings False
    SQLz = "Delete * From Schedule_Group_WorkTable;"
    DoCmd.RunSQL (SQLz)
    SQLz = "INSERT INTO Schedule_Group_WorkTable ( ScheduleId ) SELECT " & LKeyID & " AS SchId;"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub
Public Sub TreeVeiwFindIT(ByVal SearchString As String)
    Dim i      As Long
    Dim FirstI      As Long
    Call ContractAll
    Set tv = getTV(TreeForm)
    On Error GoTo GetOutLoop
    For i = 2 To DCount("UnitID", "Units") + 2000
    'For i = 2 To (tv.Nodes.Count * -1)
        If (InStr(1, tv.Nodes.Item(i).Text, SearchString)) And Not tv.Nodes.Item(i).Tag Like "*unit*" Then
            Set tv.SelectedItem = tv.Nodes.Item(i)
            If FirstI = 0 Then FirstI = i
            tv.Nodes.Item(FirstI).EnsureVisible
        End If
    Next i
    On Error GoTo 0
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
Exit Sub
GetOutLoop:
Resume Next
End Sub
Public Sub TreeVeiwFindKey(ByVal SearchString As String)
    Dim i      As Long
    Call ContractAll
     'TreeVeiwFindKey = ""
    Set tv = getTV(TreeForm)
    On Error GoTo GetOutLoop
    For i = 2 To DCount("UnitID", "Units") + 3000
    'For I = 2 To (tv.Nodes.Count * -1)
        If (InStr(1, tv.Nodes.Item(i).Key, SearchString)) Then
            Set tv.SelectedItem = tv.Nodes.Item(i)
            tv.Nodes(tv.SelectedItem.Parent.Parent.Parent.Index).EnsureVisible
            Exit For
        End If
    Next i
'cleanup
   Set NodX = Nothing
   Set tv = Nothing
   Set NodChildX = Nothing
Exit Sub
GetOutLoop:
Resume Next
End Sub
