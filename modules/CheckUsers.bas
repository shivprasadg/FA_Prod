Attribute VB_Name = "CheckUsers"
Option Compare Database
#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later
Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
Private Declare Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Function FAUserName() As String
    ' Returns the network login name
    FAUserName = LCase(Environ("username"))
End Function
Function FAUserAccess() As Integer
    ' Returns the Users' Access Level
    FAUserAccess = DLookup("[Access Level]", "TBL_Employees", "[User ID] = FAUserName()")
End Function

Sub CheckUserActive()
    If DLookup("[Active User]", "TBL_Employees", "[User ID] = FAUserName()") = True Then
        Exit Sub
    Else
        DoCmd.CloseDatabase
    End If
End Sub
Function FAUserAccessLevel() As String
    ' Returns the Users' Access Level Int
    FAUserAccessLevel = DLookup("[Access Level]", "TBL_Employees", "[User ID] = FAUserName()")
End Function
Function FAUserAccessGroup() As String
    ' Returns the Users' Access Group
    FAUserAccessGroup = DLookup("[AccessGroup]", "TBL_Employees", "[User ID] = FAUserName()")
End Function
Function FAUserFullName() As String
    ' Returns the Users' Name
    FAUserFullName = DLookup("[Employee Name]", "TBL_Employees", "[User ID] = FAUserName()")
End Function
Function FAUserIDx() As Integer
    ' Returns the Users' ID
    FAUserIDx = Trim(DLookup("[ID]", "TBL_Employees", "[User ID] = FAUserName()"))
End Function
Function FAUserInit() As String
    ' Returns the Users' ID
    FAUserInit = Trim(DLookup("empint", "TBL_Employees", "[User ID] = FAUserName()"))
End Function

Function checkgarage() As Integer
    If IsNull(DLookup("[UnitGarageLocationRef]", "Units", "[UnitID] = Forms!UnitsForm.UnitId")) Then
        checkgarage = 1
    Else
        checkgarage = 0
    End If
End Function

Function checkCostacctcd() As Integer
    If IsNull(DLookup("[UCCC]", "Units", "[UnitID] = Forms!UnitsForm.UnitId")) Then
        checkCostacctcd = 1
    Else
        checkCostacctcd = 0
    End If
End Function

Public Function Environ(Expression)
    On Error GoTo Err_Environ
    Environ = VBA.Environ(Expression)
Exit_Environ:
    Exit Function
Err_Environ:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Environ
End Function
