SELECT
  Vendors.VendorName,
  Vendors.VendorAddress,
  [VendorCity] & " " & [vendorstate] & " " & [vendorzip] AS VendorCst,
  Vendors.VendorPhone,
  Units.unitid AS [Asset No],
  "14-635" AS [PO FAID],
  IIf([UnitPilot] = True, "Yes", Null) AS [Pilot?],
  Units.Unitvin AS VIN,
  IIf(
    IsNull([locsublessee]),
    [clientcompanyname],
    [locsublessee]
  )& " " & [locationstreet] & " " & [LAddrLine1] & " " & [LCity] & " " & [LState] & " " & [LZip] AS [Deliver to],
  [unityr] & " " & [unitmake] & " " & [unitmodel] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& " " & [unittype] & " " & IIf([axle] = "TBD", Null, [axle]) AS Unit,
  Units.UnitFinalCost AS [Final Cost],
  Units.UnitBuildDate AS [Build DT],
  Units.unitestDelDate AS [Est Delivery],
  Units.UnitActualDelDate AS [Del Date],
  Units.UnitEngineSerial AS [Engine Serial#],
  Units.UnitTransSerial AS [Transmission Serial #],
  Units.UnitFASerial AS [Drive Axle Serial #],
  Units.UnitRAFSerial AS [F Rear Axle Serial #],
  Units.UnitRARSerial AS [R Rear Axle Serial #]
FROM
  (
    (
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
      LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    )
    LEFT JOIN Vendors ON Units.VendorREF = Vendors.VendorID
  )
  INNER JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
WHERE
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.BLCD) Is Null
    )
  )
  OR (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.BLCD) Is Null
    )
  )
  OR (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.BLCD) Is Null
    )
  );
