dbMemo "SQL" ="TRANSFORM First(Vendors.VendorSNm) AS ChildVendor\015\012SELECT Cost_ChildParts."
    "AssetID\015\012FROM Cost_ChildParts INNER JOIN Vendors ON Cost_ChildParts.ChildV"
    "endor = Vendors.VendorID\015\012WHERE (((Cost_ChildParts.SchID)=3056))\015\012GR"
    "OUP BY Cost_ChildParts.AssetID\015\012PIVOT [PartType] & ' Vendor';\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
