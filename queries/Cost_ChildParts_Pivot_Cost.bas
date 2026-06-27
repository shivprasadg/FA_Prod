dbMemo "SQL" ="TRANSFORM First(Cost_ChildParts.PartCost) AS FirstOfPartCost\015\012SELECT Cost_"
    "ChildParts.AssetID\015\012FROM Cost_ChildParts\015\012WHERE (((Cost_ChildParts.S"
    "chID)=3056))\015\012GROUP BY Cost_ChildParts.AssetID\015\012PIVOT [PartClass] & "
    "Chr (13) & Chr (10) & [PartType] & ' Cost';\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
