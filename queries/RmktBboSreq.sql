SELECT
  Units.UnitID,
  Rmkt.RmktId,
  Rmkt.BuyerRef,
  Clients.clientID,
  Clients.clientCompanyName,
  Clients.clientCoNm2,
  Clients.clientDBA,
  Clients.clientAddress1,
  Clients.clientAddress2,
  Clients.clientCity,
  Clients.clientState,
  Clients.clientZip AS Buyerzip,
  Clients.[Dlr/Resale],
  Units.unitvin,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  Rmkt.SoldPrice,
  Rmkt.TargetSale,
  Rmkt.InvoiceRequested,
  Units.unitstatus,
  Rmkt.RmkSm,
  Rmkt.InvoicePaidDate,
  Rmkt.InvoiceRequestNum,
  Rmkt.InvoiceRequestDate,
  Rmkt.InvoicePaidDate,
  Clients_1.clientCompanyName AS SellerCoNm,
  Clients_1.clientCoNm2 AS SellerCo2,
  Clients_1.clientDBA AS SellerDBA,
  Clients_1.clientAddress1 AS SellerAdd1,
  Clients_1.clientAddress2 AS SellerAdd2,
  Clients_1.clientCity AS SellerCity,
  Clients_1.clientState AS SellerSt,
  Clients_1.clientZip AS SellerZip,
  Rmkt.SoldPrice,
  Locations.LCity,
  Locations.LState,
  Locations.LZip,
  Rmkt.pbos,
  "<Strong>" & [clients].[clientCompanyName] & "</Strong><br>" & IIf(
    IsNull([clients].[clientCoNm2]),
    Null,
    [clients].[clientCoNm2] & "<br>"
  )& IIf(
    IsNull([clients].[clientDBA]),
    Null,
    [clients].[clientDBA] & "<br>"
  )& IIf(
    IsNull([clients].[clientAddress1]),
    Null,
    [clients].[clientAddress1] & "<br>"
  )& IIf(
    IsNull([clients].[clientAddress2]),
    Null,
    [clients].[clientAddress2] & "<br>"
  )& [clients].[clientcity] & " " & [clients].[clientState] & " " & [clients].[clientZip] AS Buyer,
  CCur(
    Nz([rmkt].[SoldPrice], 0)
  ) AS SalePrice,
  "<Strong>" & [clients_1].[clientCompanyName] & "</Strong><br>" & IIf(
    IsNull([clients_1].[clientCoNm2]),
    Null,
    [clients_1].[clientCoNm2] & "<br>"
  )& IIf(
    IsNull([clients_1].[clientDBA]),
    Null,
    [clients_1].[clientDBA] & "<br>"
  )& IIf(
    IsNull([clients_1].[clientAddress1]),
    Null,
    [clients_1].[clientAddress1] & "<br>"
  )& IIf(
    IsNull([clients_1].[clientAddress2]),
    Null,
    [clients_1].[clientAddress2] & "<br>"
  )& [clients_1].[clientcity] & " " & [clients_1].[clientstate] & " " & [clients_1].[clientZip] AS Client,
  LookUp_UnitDesc.DescriptionMed
FROM
  (
    (
      (
        (
          Clients AS Clients_1
          INNER JOIN MstrLease ON Clients_1.clientID = MstrLease.Client
        )
        INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
      )
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN (
      (
        (
          Units
          INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
        )
        INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
      )
      LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    ) ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN LookUp_UnitDesc ON Units.UnitID = LookUp_UnitDesc.AssetID
WHERE
  (
    (
      (Rmkt.InvoiceRequestNum)= forms!ReMarketInvoiceManager!InvoiceRequestID
    )
    And (
      (Rmkt.InvoicePaidDate) Is Not Null
    )
    And (
      (Rmkt.pbos)= True
    )
  )
ORDER BY
  Units.unitvin;
