SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitVIN,
  Vendors.VendorName,
  Vendors.VendorAddress,
  Vendors.VendorCity,
  Vendors.VendorState,
  Vendors.VendorZip,
  ChildParts.ChildInvoiceNum,
  Vendors.VendorPmtType,
  ChildParts.childFinalCost,
  Vendors.VENDORPAYEE,
  Vendors.VENDORABA,
  Vendors.VENDORACH,
  Vendors.VENDORBKACCT,
  Vendors.VENDORBANK,
  Vendors.VENDORBADD1,
  Vendors.VEDNORBADD2,
  Vendors.VENDORBCITY,
  Vendors.VENDORBST,
  Vendors.VENDORBZIP,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.ChildVIN
FROM
  (
    (
      (
        (
          Clients
          INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
        )
        INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
      )
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN (
    ChildParts
    LEFT JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID
  ) ON Units.UnitID = ChildParts.UnitREF
WHERE
  (
    (
      (MstrLease.MLNo)= [enter Master Lease No]
    )
    AND (
      (ChildParts.ChildInvoiceNum)= [enter invoice no]
    )
  );
