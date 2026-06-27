dbMemo "SQL" ="TRANSFORM First(Cost_ChildParts.PartInvoice) AS FirstOfPartInvoice\015\012SELECT"
    " Cost_ChildParts.AssetID\015\012FROM Cost_ChildParts\015\012WHERE (((Cost_ChildP"
    "arts.SchID)=3056))\015\012GROUP BY Cost_ChildParts.AssetID\015\012PIVOT [PartCla"
    "ss] & Chr (13) & Chr (10) & [PartType] & ' Invoice';\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
