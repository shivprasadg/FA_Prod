SELECT
  Vendors.VendorName,
  Vendors.VendorAddress,
  [VendorCity] & " " & [vendorstate] & " " & [vendorzip] AS VendorCst,
  Vendors.VendorPhone,
  Count(Units.UnitID) AS CountOfUnitID,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [UnitSubtype] & " " & [UnitType] & IIf([axle] = "TBD", Null, " " & [axle]) AS Unit,
  Clients.clientCompanyName AS CompanyName,
  Min(Units.UnitActualDelDate) AS MinDeliveryDate,
  Max(Units.UnitActualDelDate) AS MaxDeliveryDate,
  Schedule.schfaid AS FAID,
  Locations.LState,
  Locations.LCity,
  Locations.LAddrLine1,
  Locations.LocationStreet,
  Locations.LZip,
  Units.UnitFinalCost,
  Locations.LocationID,
  Min(Units.UnitEstDelDate) AS MinOfUnitEstDelDate,
  Max(Units.Unitestdeldate) AS MaxOfUnitestdeldate,
  Locations.LocSubLessee,
  1 AS rptsec
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
GROUP BY
  Vendors.VendorName,
  Vendors.VendorAddress,
  [VendorCity] & " " & [vendorstate] & " " & [vendorzip],
  Vendors.VendorPhone,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [UnitSubtype] & " " & [UnitType] & IIf([axle] = "TBD", Null, " " & [axle]),
  Clients.clientCompanyName,
  Schedule.schfaid,
  Locations.LState,
  Locations.LCity,
  Locations.LAddrLine1,
  Locations.LocationStreet,
  Locations.LZip,
  Units.UnitFinalCost,
  Locations.LocationID,
  Locations.LocSubLessee,
  1,
  MstrLease.MLOrig,
  SchGrp.BLCD,
  SchGrp.sgrpfaid,
  Units.unitfaid
HAVING
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.BLCD) Is Null
    )
  )
ORDER BY
  Locations.LState,
  Locations.LCity,
  Locations.LocationStreet,
  Units.UnitFinalCost;
