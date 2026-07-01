Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Enter_System_Click()
    If FaUserName = DLookup("[User ID]", "TBL_Employees", "[User ID] = FAUserName()") And DLookup("[Active User]", "TBL_Employees", "[User ID] = FAUserName()") = True Then
        DoCmd.OpenForm "FAMainMenu", acNormal, "", "", , acNormal
        If FAUserAccessLevel = 3 Then DoCmd.OpenForm "AcctMenu", acNormal, "", "", , acNormal
        If FAUserAccessLevel = 4 Then DoCmd.OpenForm "AM_Menu", acNormal, "", "", , acNormal
        If FAUserAccessLevel = 5 Then DoCmd.OpenForm "DocsMenu", acNormal, "", "", , acNormal
        If FAUserAccessLevel = 6 Then DoCmd.OpenForm "ManagementMenu", acNormal, "", "", , acNormal
        If FAUserAccessLevel = 7 Then DoCmd.OpenForm "OPS VP Menu", acNormal, "", "", , acNormal
        If FAUserAccessLevel = 8 Then DoCmd.OpenForm "Marketing VP Menu", acNormal, "", "", , acNormal
'       If FAUserAccessLevel = 9 Then DoCmd.OpenForm "FAMainMenu"
        If FAUserAccessLevel = 10 Then DoCmd.OpenForm "Rmkt Menu"
        If FAUserAccessLevel = 11 Then DoCmd.OpenForm "ManagementMenuPF"
        DoCmd.OpenForm "USERACTIVECHECK", acNormal, "", "", , acHidden
4        Call LogUsersin
    Else
        Beep
        MsgBox "You are not authorized to use the Fleet Advantage Data Base at this time. Please contact technical support to gain access", vbCritical, "LOGIN"
        DoCmd.CloseDatabase
        Exit Sub

    End If
    DoCmd.OpenForm "FAAMenu"
End Sub



Private Sub Form_Load()
DoCmd.OpenForm "USERACTIVECHECK", acNormal, "", "", , acHidden
Call Enter_System_Click
End Sub
