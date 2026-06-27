Attribute VB_Name = "Userlog"
Option Compare Database
Option Explicit

Sub LogUsersin()
    'Dim tbl_UserlogIN As DAO.Recordset
    Dim UserlogDet As DAO.Recordset

    Set UserlogDet = CurrentDb.OpenRecordset("tbl_UserlogIN", dbOpenDynaset, dbSeeChanges)
    UserlogDet.AddNew
    UserlogDet("Time_In_Stamp").Value = Now()
    UserlogDet("User_Name").Value = FAUserName
    UserlogDet.Update
End Sub

Sub LogUsersout()
    'Dim tbl_UserlogOUT As DAO.Recordset
    Dim UserlogDet As DAO.Recordset
    Set UserlogDet = CurrentDb.OpenRecordset("tbl_UserlogOUT", dbOpenDynaset, dbSeeChanges)
    UserlogDet.AddNew
    UserlogDet("Time_OUT_Stamp").Value = Now()
    UserlogDet("User_Name").Value = FAUserName
    UserlogDet.Update
End Sub
