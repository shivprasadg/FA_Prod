Operation =6
Option =0
Where ="(((Cost_ChildParts.SchID)=3015))"
Begin InputTables
    Name ="Cost_ChildParts"
End
Begin OutputColumns
    Expression ="Cost_ChildParts.AssetID"
    GroupLevel =2
    Expression ="[PartMake] & Chr(13) & Chr(10) & [PartType] & ' Cost'"
    GroupLevel =1
    Alias ="FirstOfPartCost"
    Expression ="First(Cost_ChildParts.PartCost)"
End
Begin Groups
    Expression ="Cost_ChildParts.AssetID"
    GroupLevel =2
    Expression ="[PartMake] & Chr(13) & Chr(10) & [PartType] & ' Cost'"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =1449
    Bottom =1337
    Left =-1
    Top =-1
    Right =1425
    Bottom =1017
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Cost_ChildParts"
        Name =""
    End
End
