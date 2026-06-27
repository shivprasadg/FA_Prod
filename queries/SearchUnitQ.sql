SELECT
  NewULookupQ.UnitID,
  NewULookupQ.CompanyNm,
  NewULookupQ.ClientShNm,
  NewULookupQ.VendorSNm,
  NewULookupQ.SalespersonName,
  NewULookupQ.CompanyNm,
  NewULookupQ.SPInitials,
  NewULookupQ.ClientShNm,
  NewULookupQ.VendorREF,
  NewULookupQ.UnitVIN,
  NewULookupQ.UnitPlateNum,
  NewULookupQ.UnitPlateNumH,
  NewULookupQ.UnitUnitNum,
  NewULookupQ.UnitYr,
  NewULookupQ.UnitMake,
  NewULookupQ.UnitModel,
  NewULookupQ.UnitType,
  NewULookupQ.UnitVendorPytDate,
  NewULookupQ.clientState,
  NewULookupQ.clientCity,
  "Edit" AS Edit,
  NewULookupQ.SGrpID,
  NewULookupQ.GroupR,
  NewULookupQ.UnitStatus,
  NewULookupQ.LCity,
  NewULookupQ.LState,
  Clients.ClientShNm AS RBuyer,
  Schedule.SchNo,
  NewULookupQ.UnitInServ,
  NewULookupQ.unitinservind
FROM
  Schedule
  INNER JOIN (
    (
      (
        NewULookupQ
        LEFT JOIN Rmkt ON NewULookupQ.UnitID = Rmkt.UnitRef
      )
      LEFT JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
    )
    INNER JOIN SchGrp ON NewULookupQ.SGrpID = SchGrp.SGrpID
  ) ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (NewULookupQ.ClientGroupID)<> 18
    )
  );
