SELECT
  [vendorref] & [childvendorref] & [locationid] & [unitfinalcost] & [childfinalcost] AS [KEY],
  Units.UnitID,
  Vendors.VendorName,
  Vendors.VendorID,
  Vendors.VendorSNm,
  Vendors.VendorAddress,
  Vendors.VendorCity,
  Vendors.VendorState,
  Vendors.VendorZip,
  Vendors_1.vendorid,
  Vendors_1.VendorName,
  Vendors_1.VendorState,
  Vendors_1.VendorZip,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  Units.UnitEstDelDate,
  Units.unitfinalcost,
  ChildParts.childyear,
  ChildParts.childmake,
  ChildParts.ChildClass,
  ChildParts.childtype,
  ChildParts.childfinalcost,
  IIf(
    IsNull([unitfaid]),
    IIf(
      IsNull([sgrpfaid]),
      [schfaid],
      [sgrpfaid]
    ),
    [unitfaid]
  ) AS FAID,
  Schedule.SchFAID,
  SchGrp.SGrpFAID,
  Units.UnitFAID,
  MstrLease.MLOrig
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN (
    SchGrp
    INNER JOIN (
      (
        (
          (
            Units
            LEFT JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
          )
          INNER JOIN Vendors ON Units.VendorREF = Vendors.VendorID
        )
        LEFT JOIN Vendors AS Vendors_1 ON ChildParts.ChildVendorREF = Vendors_1.VendorID
      )
      LEFT JOIN Locations ON Units.UnitDeliveryLocationRef = Locations.LocationID
    ) ON SchGrp.SGrpID = Units.GroupREF
  ) ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (
        IIf(
          IsNull([unitfaid]),
          IIf(
            IsNull([sgrpfaid]),
            [schfaid],
            [sgrpfaid]
          ),
          [unitfaid]
        )
      )= "13-547"
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  );
