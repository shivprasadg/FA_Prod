Attribute VB_Name = "LOMenuLog"
Option Compare Database

Sub Lomenuloger()
 Dim FADB As DAO.DataBase
    Dim lomenulog As DAO.Recordset
    Dim UserlogDet As DAO.Recordset

    Set FADB = CurrentDb
    Set UserlogDet = FADB.OpenRecordset("LOMenuLog", dbOpenDynaset, dbSeeChanges)

    UserlogDet.AddNew
    UserlogDet("Date").Value = Now()
    UserlogDet("UserName").Value = FAUserName
    UserlogDet.Update
End Sub


Sub Loreportloger()
 Dim FADB As DAO.DataBase
    Dim TargetProspectformlog As DAO.Recordset
    Dim UserlogDet As DAO.Recordset

    Set FADB = CurrentDb
    Set UserlogDet = FADB.OpenRecordset("TargetProspectformlog", dbOpenDynaset, dbSeeChanges)

    UserlogDet.AddNew
    UserlogDet("Date").Value = Now()
    UserlogDet("UserName").Value = FAUserName
    UserlogDet.Update
End Sub
