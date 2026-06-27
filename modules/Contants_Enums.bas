Attribute VB_Name = "Contants_Enums"
Option Compare Database

Public Enum eMouse
    LeftClick = 1
    Rightclick = 2
End Enum

Public Function AtlaasVersion() As String
 AtlaasVersion = CurrentDb.Properties("AppTitle")
End Function
