dbMemo "SQL" ="TRANSFORM First(Cost_ChildParts.PartPaidDate) AS FirstOfPartPaidDate\015\012SELE"
    "CT Cost_ChildParts.AssetID\015\012FROM Cost_ChildParts\015\012WHERE (((Cost_Chil"
    "dParts.SchID)=3056))\015\012GROUP BY Cost_ChildParts.AssetID\015\012PIVOT [PartC"
    "lass] & Chr (13) & Chr (10) & [PartType] & ' Paid';\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
